import bpy
import logging
import os

# === LOGGER SETUP HELPERS ===
def setup_logger(name, filename=None):
    """Create and return a logger with an optional filename."""
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:
        log_path = bpy.path.abspath(f"//{filename or f'{name}.log'}")
        handler = logging.FileHandler(log_path, mode='w')
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger

# === STRIP UTILITIES ===
def deselect_all():
    """Deselect all strips in the sequence editor."""
    scene = bpy.context.scene
    seq = scene.sequence_editor

    if not seq:
        print("No sequence editor found.")
        return

    for s in seq.sequences_all:
        s.select = False

def select_strip(strip):
    seq = get_seq()
    deselect_all()
    strip.select = True
    seq.active_strip = strip

def get_bounds(strip):
    """Return the start and end frames of a strip."""
    return int(strip.frame_final_start), int(strip.frame_final_end)

def get_seq():
    """Return the sequence editor object."""
    scene = bpy.context.scene
    seq = scene.sequence_editor

    if not seq:
        raise Exception("No sequence editor found. Aborting.")
    return seq


def select_strips_intersecting_frame(seq, frame):
    """Select all strips that intersect the given frame."""
    selected = []
    for strip in seq.sequences_all:
        s_start, s_end = get_bounds(strip)
        strip.select = s_start < frame < s_end
        if strip.select:
            selected.append(strip.name)
    return selected

def should_speed_up(strip):
    """Check if the strip should have speed effect applied."""
    return strip.type == 'MOVIE'

def get_markers():
    """Return all timeline markers."""
    return bpy.context.scene.timeline_markers

def get_sorted_markers():
    seq = get_seq()
    #markers = sorted([m for m in get_markers()])
    get_frame = lambda m : m.frame
    sorted_markers = sorted([m for m in get_markers()], key=get_frame)
    return sorted_markers

def set_marker(marker, frame):
    marker.select = True
    marker.frame = frame
    marker.select = False
    return marker.frame

# TODO: rename, colorize and also shift markers!

# === CUTTING STRIPS AT MARKERS ===
def cut_strips_at_markers():
    """Split all strips at every timeline marker."""
    logger = setup_logger("cut_logger", "cut_at_markers_log.txt")
    seq = get_seq()
    markers = get_sorted_markers()
    logger.info(f"Found {len(markers)} markers for cutting: {markers}")

    for i,m in enumerate(markers):
        frame = m.frame
        bpy.context.scene.frame_current = frame
        logger.info(f"Cutting at frame {frame}")
        selected_strips = select_strips_intersecting_frame(seq, frame)
        logger.debug(f"Selected {len(selected_strips)} strips: {selected_strips}")
        bpy.ops.sequencer.split(frame=frame, type='SOFT', side='BOTH')

        # TODO: check if working
        m.select = True
        i_clip = int(i / 2)
        new_name = f"{i_clip}_{m.name}"
        if (i % 2) == 0:
            new_name = "start" + new_name
            #m.color = 'GREEN'
        elif (i % 2) == 1:
            new_name = "stop" + new_name
            #m.color = 'RED'
        m.name = new_name
        m.select = False


    logger.info(f"✅ Completed cutting at {len(markers)} markers.")

# === SPEED UP EVEN INTERVALS ===
def speed_up_sections_between_even_markers(speed_factor=10):
    """Apply speed effect to strips between even-indexed marker intervals."""
    logger = setup_logger("speed_logger", "speed_up_log.txt")
    seq = get_seq()

    markers = sorted(get_markers(), key=lambda m: m.frame)
    logger.info(f"📌 Found {len(markers)} markers: {[m.frame for m in markers]}")

    if len(markers) < 4:
        logger.warning("⚠️ Not enough markers (need at least 4). Exiting.")
        return

    area_type = 'SEQUENCE_EDITOR'
    areas = [a for a in bpy.context.window.screen.areas if a.type == area_type]
    if not areas:
        logger.error("No SEQUENCE_EDITOR area open. Aborting.")
        return

    def trim_speed(strip):
        start, end = get_bounds(strip)
        duration = max(30, int((end - start) / speed_factor))
        strip.frame_final_end = start + duration
        logger.debug(f"⏱️ Trimmed '{strip.name}' to {duration} frames.")

    def squeeze_strip(strip):
        deselect_all()
        strip.select = True
        seq.active_strip = strip
        start, end = get_bounds(strip)

        try:
            bpy.ops.sequencer.effect_strip_add(
                type='SPEED',
                frame_start=start,
                frame_end=end,
                channel=strip.channel + 1,
                replace_sel=False
            )
            speed_strip = seq.sequences[-1]
            speed_strip.speed_factor = speed_factor

            #trim_speed(strip)
            start, end = get_bounds(strip)
            duration = max(30, int((end - start) / speed_factor))
            strip.frame_final_end = start + duration
            logger.debug(f"⏱️ Trimmed '{strip.name}' to {duration} frames.")

            logger.info(f"✅ Applied speed to '{strip.name}' ({start}-{end})")
        except Exception as e:
            logger.error(f"❌ Failed on '{strip.name}': {e}")

    with bpy.context.temp_override(
        window=bpy.context.window,
        area=areas[0],
        region=[r for r in areas[0].regions if r.type == 'WINDOW'][0],
        screen=bpy.context.window.screen
    ):
        


        for i in range(1, len(markers) - 1, 2):
            m_start, m_end = markers[i].frame, markers[i + 1].frame
            if m_end <= m_start:
                logger.warning(f"⚠️ Skipping invalid interval {m_start}–{m_end}")
                continue


            markers_to_shift = []

            logger.info(f"🟦 Speeding interval {i // 2}: {m_start} → {m_end}")
            for strip in list(seq.sequences_all):
                m_start, m_end = markers[i].frame, markers[i + 1].frame
                s_start, s_end = get_bounds(strip)
                if s_start >= m_start and s_end <= m_end:
                    if should_speed_up(strip):
                        
                        #squeeze_strip(strip)
                        deselect_all()
                        strip.select = True
                        seq.active_strip = strip
                        s_start, s_end = get_bounds(strip)

                        try:
                            bpy.ops.sequencer.effect_strip_add(
                                type='SPEED',
                                frame_start=s_start,
                                frame_end=s_end,
                                channel=strip.channel + 1,
                                replace_sel=False
                            )
                            speed_strip = seq.sequences[-1]
                            speed_strip.speed_factor = speed_factor

                            #trim_speed(strip)
                            new_duration = max(30, int((s_end - s_start) / speed_factor))
                            new_s_end = s_start + new_duration
                            strip.frame_final_end = new_s_end
                            # TODO: check whether marker needs to be moved as well!
                            # when you are at a cut - at a marker
                            if s_end == m_end:
                                # marker needs to be shifted as well ...
                                #set_marker(markers[i+1], new_s_end)
                                markers_to_shift.append([markers[i+1], new_s_end])
                                logger.debug(f"⏱️ marker '{markers[i+1].name}' to shift along with strip {strip.name} to frame {new_s_end}.")
                                

                            logger.debug(f"⏱️ Trimmed '{strip.name}' to {new_duration} frames.")

                            logger.info(f"✅ Applied speed to '{strip.name}' ({s_start}-{s_end})")
                            s_start, s_end = get_bounds(strip)
                            logger.info(f"✅ New interval for '{strip.name}' ({s_start}-{s_end})")


                        except Exception as e:
                            logger.error(f"❌ Failed on '{strip.name}': {e}")

                        # TODO: check if working
                        #s_start_new, s_end_new = get_bounds(strip)
                        #set_marker(markers[i], s_start_new)
                        #set_marker(markers[i+1], s_end_new)

                    elif strip.type == 'SOUND':
                        logger.info(f"🗑️ Deleting audio '{strip.name}'")
                        try:
                            seq.sequences.remove(strip)
                        except Exception as e:
                            logger.error(f"❌ Failed to delete '{strip.name}': {e}")

            # shift markers
            for m, _new_frame in markers_to_shift:
                set_marker(m, _new_frame)
                logger.debug(f"⏱️ marker '{m.name}' shifted to frame {m.frame}.")

        # TODO: move all markers

# === REMOVE GAPS BETWEEN STRIPS ===
def remove_channel_strip_gaps():
    """Remove frame gaps between grouped strip blocks on the timeline."""
    logger = setup_logger("gap_logger", "strip_gap_removal_log.txt")
    seq = get_seq()
    strips = sorted(seq.sequences_all, key=lambda s: s.frame_final_start)

    def group_strips(strips):

        # Collect and sort all strips based on their start frame
        strips = [s for s in get_seq().sequences_all]
        strips.sort(key=lambda s: s.frame_final_start)

        # Group strips by "channel blocks" – meaning strips that start and end together
        groups = []

        # count the group index
        group_i = 0
        _s = strips[group_i]
        first_channel_strip = {
            "type": "strip",
            "frame_final_start": _s.frame_final_start,
            "frame_final_end": _s.frame_final_end,
            "strips": []
        }
        groups.append(first_channel_strip)

        # Assign strips to channel strip groups
        for i_s, s in enumerate(strips):
            if s.frame_final_start == groups[group_i]["frame_final_start"]:
                if s.frame_final_end == groups[group_i]["frame_final_end"]:
                    groups[group_i]["strips"].append(s)
                    logger.debug(f"Added strip {s.name} ({s.type}) to channel group {group_i}")
            elif s.frame_final_start < groups[group_i]["frame_final_start"]:
                raise Exception(
                    f"❌ Overlap error: Missed a strip: {s.name}, {s.type}, "
                    f"({s.frame_final_start}-{s.frame_final_end}) vs. current group "
                    f"({groups[group_i]['frame_final_start']}-{groups[group_i]['frame_final_end']})"
                )
            elif s.frame_final_start >= groups[group_i]["frame_final_end"]:
                # Start a new group when a new, non-overlapping strip is found
                new_group = {
                    "type": "strip",
                    "frame_final_start": s.frame_final_start,
                    "frame_final_end": s.frame_final_end,
                    "strips": [s]
                }
                groups.append(new_group)
                group_i += 1
                logger.debug(f"Started new channel group {group_i} with strip {s.name} ({s.type})")


        # TODO: add markers as well!
        sorted_markers = get_sorted_markers()
        _m = sorted_markers[0]
        first_marker_strip = {
            "type": "marker",
            "frame": _m.frame,
            "marker": _m
        }
        for _m in sorted_markers:
            new_marker_group = {
                "type": "marker",
                "frame": _m.frame,
                "marker": _m
            }
            groups.append(new_marker_group)
        # TODO: sort group list
        def get_group_start(group):
            if group["type"] == "marker":
                # this is such that markers will always be moved first
                # as they will not result in shifting cursor/current_frame
                return float(group["frame"])-0.1
            elif group["type"] == "strip":
                return float(group["frame_final_start"])
            else:
                raise Exception("unknown group type!")
                #return -1   
        sorted_groups = sorted(groups, key=get_group_start)
        #sorted(groups, key=get_group_start)

        group_frames = [get_group_start(_g) for _g in sorted_groups]
        group_types = [_g["type"] for _g in sorted_groups]
        logger.debug(f" Sorted groups in ascending frames: {tuple(zip(group_frames, group_types))}")

        logger.debug(f"➡️ Found {len(groups)} groups")

        #return groups
        return sorted_groups

    #sorted_markers = get_sorted_markers()

    grouped_strips = group_strips(strips)

    current_frame = 0
    for i,group in enumerate(grouped_strips):
        offset = 0
        logger.debug(f"➡️ Current frame {current_frame}")
        logger.debug(f"➡️ Current group {group}")
        # TODO: include markers

        if group["type"] == "marker":
            if group["frame"] < current_frame:
                logger.debug(f"❌ Ignoring marker {group["marker"]} with frame {group['frame']} that comes before current_frame={current_frame}")
            
            # only shift if group frame is ahead of current frame!
            elif group["frame"] == current_frame:
                pass
            elif group["frame"] > current_frame:
                offset = group["frame"] - current_frame
                _m = group["marker"]
                old_frame = _m.frame
                _m.frame -= offset
                logger.debug(f"➡️ Shifted {_m.name} by {-offset} frames from {old_frame} to {_m.frame}")
        
            # marker does not shift current frame as it has no length along timeline
            #current_frame = group["frame"] - offset
            #current_frame = current_frame

        elif group["type"] == "strip":
            if group["frame_final_start"] < current_frame:
                logger.debug(f"❌ Ignoring group {[s.name for s in group['strips']]} with frame {group['frame_final_start']} that comes before current_frame={current_frame}")
                #raise Exception(f"❌ Overlap in {[s.name for s in group['strips']]} at {group['frame_final_start']}")
            
            # only shift if group frame is ahead of current frame!
            elif group["frame_final_start"] == current_frame:
                pass
            elif group["frame_final_start"] > current_frame:
                offset = group["frame_final_start"] - current_frame
                for s in group["strips"]:
                    if s.type != "SPEED":
                        s.frame_start -= offset
                        # TODO: end needs to be adjustes as well? - THE BELOW SEEMS TO BE BUGGY!!!
                        #s.frame_final_start -= offset
                        #s.frame_final_end -= offset
                        logger.debug(f"➡️ Shifted {s.name} by {-offset} frames to {get_bounds(s)}")
        
            current_frame = group["frame_final_end"] - offset



            # TODO: check if working!
            # TODO: shift markers in sync with strips to adapt gaps to marker positions as well
            # NOPE DOESNT WORK: CAN NOT EXPECT THAT MARKERS ALWAYS COINCIDE WITH END OF STRIPS; THEY MAY LIE IN GAPS
            #for i_m in range(i_m_current, len(sorted_markers)):
            #    if group["frame_final_start"] == sorted_markers[i_m].frame:
            #        set_marker(sorted_markers[i_m], group["frame_final_start"] - offset)
            #        logger.debug(f"➡️ Shifted marker {sorted_markers[i_m].name} by {-offset} frames to {sorted_markers[i_m].frame}")
            #        i_m_current = i_m
            #        continue
            #for i_m in range(i_m_current, len(sorted_markers)):
            #    if group["frame_final_end"] == sorted_markers[i_m].frame:
            #        set_marker(sorted_markers[i_m], group["frame_final_end"] - offset)
            #        logger.debug(f"➡️ Shifted marker {sorted_markers[i_m].name} by {-offset} frames to {sorted_markers[i_m].frame}")
            #        i_m_current = i_m
            #        continue

            # NOPE DOESNT WORK: 
            # Every time a group is shifted ...
            # every other strip and marker needs to be shifted accordingly
            #for s in strips:
            #    s_start, s_end = get_bounds(s)
            #    if s.type != "SPEED":
            #        if s_start >= group["frame_final_end"]:
            #            # shift ALL ahead markers by the same amount
            #            s.frame_final_start -= offset
            #            s.frame_final_end -= offset
            #            #set_marker(sorted_markers[i_m], sorted_markers[i_m].frame - offset)
            #            logger.debug(f"➡️ Shifted strip {s.name} by {-offset} from {(s_start, s_end)} to {get_bounds(s)}")
            #for m in sorted_markers:
            #    if m.frame >= group["frame_final_start"]:
            #        # shift ALL ahead markers by the same amount
            #        m.frame -= offset
            #        #set_marker(sorted_markers[i_m], sorted_markers[i_m].frame - offset)
            #        logger.debug(f"➡️ Shifted marker {m.name} by {-offset} frames to {m.frame}")


        
        logger.debug(f"➡️ moving to next group {i+1} - updating current frame {current_frame}")








# === MAIN ENTRYPOINT ===
def main():
    deselect_all()
    cut_strips_at_markers()
    deselect_all()
    speed = 30
    speed_up_sections_between_even_markers(speed_factor=speed)
    deselect_all()
    remove_channel_strip_gaps()
    deselect_all()
    print("✅ All operations completed. Logs saved to the .blend directory.")





# run the whole program
main()
