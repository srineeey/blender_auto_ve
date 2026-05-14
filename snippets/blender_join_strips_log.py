import bpy
import logging

# Set up logger for debugging output
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# Define the path to the log file (relative to the .blend file)
log_path = bpy.path.abspath("//strip_gap_removal_log.txt")

# Create file handler and configure it
log_handler = logging.FileHandler(log_path)
log_handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
log_handler.setFormatter(formatter)

# Ensure the handler is not added multiple times
if not logger.handlers:
    logger.addHandler(log_handler)


def remove_channel_strip_gaps():
    scene = bpy.context.scene
    seq_editor = scene.sequence_editor

    if not seq_editor:
        logger.error("No sequence editor found.")
        return

    # Collect and sort all strips based on their start frame
    strips = [s for s in seq_editor.sequences_all]
    strips.sort(key=lambda s: s.frame_final_start)

    # Group strips by "channel blocks" – meaning strips that start and end together
    channel_strips = []
    i_ch_s = 0

    _s = strips[i_ch_s]
    first_channel_strip = {
        "frame_final_start": _s.frame_final_start,
        "frame_final_end": _s.frame_final_end,
        "strips": []
    }
    channel_strips.append(first_channel_strip)

    # Assign strips to channel strip groups
    for s in strips:
        if s.frame_final_start == channel_strips[i_ch_s]["frame_final_start"]:
            if s.frame_final_end == channel_strips[i_ch_s]["frame_final_end"]:
                channel_strips[i_ch_s]["strips"].append(s)
                logger.debug(f"Added strip {s.name} ({s.type}) to channel group {i_ch_s}")
        elif s.frame_final_start < channel_strips[i_ch_s]["frame_final_start"]:
            raise Exception(
                f"❌ Overlap error: Missed a strip: {s.name}, {s.type}, "
                f"({s.frame_final_start}-{s.frame_final_end}) vs. current group "
                f"({channel_strips[i_ch_s]['frame_final_start']}-{channel_strips[i_ch_s]['frame_final_end']})"
            )
        elif s.frame_final_start >= channel_strips[i_ch_s]["frame_final_end"]:
            # Start a new group when a new, non-overlapping strip is found
            new_ch_s = {
                "frame_final_start": s.frame_final_start,
                "frame_final_end": s.frame_final_end,
                "strips": [s]
            }
            channel_strips.append(new_ch_s)
            i_ch_s += 1
            logger.debug(f"Started new channel group {i_ch_s} with strip {s.name} ({s.type})")

    current_frame = 0  # Tracks where the next strip group should be placed

    # Iterate over each group of strips and shift them to close gaps
    for ch_s in channel_strips:
        ch_s_start = ch_s["frame_final_start"]
        ch_s_end = ch_s["frame_final_end"]

        offset = 0
        if ch_s_start < current_frame:
            raise Exception(
                f"❌ Missed overlap: {[s.name for s in ch_s['strips']]}, "
                f"({ch_s_start}-{ch_s_end}) vs. current_frame ({current_frame})"
            )
        elif ch_s_start > current_frame:
            # If there’s a gap, calculate how much to shift
            offset = ch_s_start - current_frame
            for _s in ch_s["strips"]:
                # SPEED strips are dependent on their source and shouldn't be moved independently
                if _s.type != "SPEED":
                    _s.frame_start = _s.frame_start - offset
                    logger.debug(
                        f"➡️ Shifted strip {_s.name} ({_s.type}) by {-offset} frames to new start {_s.frame_start}"
                    )

        # Update the frame position where the next block should start
        current_frame = ch_s_end - offset


remove_channel_strip_gaps()
logger.info("✅ Finished removing gaps between strips.")
