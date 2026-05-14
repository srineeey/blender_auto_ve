import bpy
import logging
import os
import datetime

# Setup logging system
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# Prevent duplicate handlers if script is re-run in the same Blender session
if not logger.handlers:
    log_path = bpy.path.abspath("//my-operator-log.txt")
    log_handler = logging.FileHandler(log_path, mode='w')  # Overwrite log on each run
    log_handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    log_handler.setFormatter(formatter)
    logger.addHandler(log_handler)

logger.info("🚀 Starting speed-up script...")

# Speed-up factor (e.g., 30x faster playback)
speed_factor = 30

# Ensure an open SEQUENCE_EDITOR area is present (required for many bpy.ops.sequencer.* calls)
area_type = 'SEQUENCE_EDITOR'
areas = [area for area in bpy.context.window.screen.areas if area.type == area_type]
if not areas:
    logger.error(f"No area of type {area_type} is open. Aborting.")
    raise Exception(f"No area of type {area_type} is open.")

# Use a temporary override context so operators like sequencer.effect_strip_add work properly
with bpy.context.temp_override(
    window=bpy.context.window,
    area=areas[0],
    region=[region for region in areas[0].regions if region.type == 'WINDOW'][0],
    screen=bpy.context.window.screen
):

    def get_strip_bounds(strip):
        """Return the actual start and end frames of the strip, using final calculated bounds."""
        strip_frame_start = int(strip.frame_final_start)
        strip_frame_end = int(strip.frame_final_end)
        return [strip_frame_start, strip_frame_end]

    def trim_strip_for_speed(strip, speed_factor):
        """
        Adjust the strip length to match the speed-up effect.
        E.g., a 90-frame strip with speed_factor=30 will become 3 frames long.
        """
        old_start, old_end = get_strip_bounds(strip)
        old_duration = old_end - old_start

        # Avoid zero-length strips
        new_duration = max(1, int(old_duration / speed_factor))
        strip.frame_final_end = old_start + new_duration

        logger.debug(f"⏱️ Trimmed strip '{strip.name}' from {old_duration} → {new_duration} frames (start={old_start})")
        return new_duration

    def apply_speed_effect_to_strip(strip, speed_factor):
        """
        Apply a SPEED effect to a movie strip.
        Also adjusts the original strip length to match the new playback duration.
        """
        scene = bpy.context.scene
        seq = scene.sequence_editor

        # Ensure only this strip is selected
        for s in seq.sequences_all:
            s.select = False
        strip.select = True
        scene.sequence_editor.active_strip = strip

        strip_frame_start, strip_frame_end = get_strip_bounds(strip)

        try:
            # Add SPEED effect in the channel above
            bpy.ops.sequencer.effect_strip_add(
                type='SPEED',
                frame_start=strip_frame_start,
                frame_end=strip_frame_end,
                channel=int(strip.channel) + 1,
                replace_sel=False
            )
            speed_effect = scene.sequence_editor.sequences[-1]
            speed_effect.speed_factor = speed_factor

            logger.info(f"✅ Applied {speed_factor}x speed effect to '{strip.name}' with bounds ({strip_frame_start}–{strip_frame_end})")

            # Trim original strip to match the new shortened duration
            trim_strip_for_speed(strip, speed_factor)

        except Exception as e:
            logger.error(f"❌ Failed to apply speed to '{strip.name}': {str(e)}")

    def speed_up_sections_between_even_markers():
        """
        Iterates over all timeline markers and applies speed-up only to intervals between
        markers at index 1 and 2, 3 and 4, 5 and 6, etc. (i.e., even-indexed *pairs* starting from index 1).
        """
        scene = bpy.context.scene
        seq = scene.sequence_editor

        markers = sorted(scene.timeline_markers, key=lambda m: m.frame)
        logger.info(f"📌 Found {len(markers)} markers: {[m.frame for m in markers]}")

        if len(markers) < 4:
            logger.warning("⚠️ Not enough markers (need at least 4). Exiting.")
            return

        for i in range(1, len(markers) - 1, 2):
            start_frame = int(markers[i].frame)
            try:
                end_frame = int(markers[i + 1].frame)
            except IndexError:
                logger.warning(f"⚠️ Skipping invalid markers index {i+1}")
                continue

            duration = end_frame - start_frame
            if duration <= 0:
                logger.warning(f"⚠️ Skipping invalid interval {start_frame}–{end_frame} (Δ={duration})")
                continue

            logger.info(f"\n🟦 Interval {i//2}: {start_frame} → {end_frame} (Δ={duration})")

            # Iterate over a copy of the strip list to allow safe modification
            for strip in list(seq.sequences_all):
                strip_frame_start, strip_frame_end = get_strip_bounds(strip)

                # Select strips fully within the current interval
                if strip_frame_start >= start_frame and strip_frame_end <= end_frame:
                    logger.debug(f"🎞️ Strip '{strip.name}' of type {strip.type} with bounds ({strip_frame_start}–{strip_frame_end})")
                    if strip.type == 'MOVIE':
                        apply_speed_effect_to_strip(strip, speed_factor)
                    elif strip.type == 'SOUND':
                        logger.info(f"🗑️ Deleting audio strip '{strip.name}' with bounds {strip_frame_start}–{strip_frame_end}")
                        try:
                            seq.sequences.remove(strip)
                        except Exception as e:
                            logger.error(f"❌ Could not delete audio strip '{strip.name}': {e}")
                else:
                    # Skip strips that overlap or lie outside the bounds
                    pass

    # Main execution starts here
    speed_up_sections_between_even_markers()

# Finalize logging
logger.info("✅ Script complete.")
logger.info(f"📄 Log saved to: {bpy.path.abspath('//my-operator-log.txt')}")

# Cleanup log handlers so script can be re-run cleanly in Blender
for handler in logger.handlers:
    handler.flush()
    handler.close()
    logger.removeHandler(handler)

# Useful debugging tips:
# scene = bpy.context.scene
# seq = scene.sequence_editor
# strips = [s for s in seq.sequences_all]
# strip_attr = {attr : getattr(strip, attr) for attr in ["name", "frame_start", "frame_final_start", "frame_offset_start", "frame_final_duration", "frame_final_end", "frame_offset_end", "frame_duration"]}
