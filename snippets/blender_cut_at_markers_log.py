import bpy
import logging

# Setup logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# Create a file handler to store logs in the blend file's directory
log_path = bpy.path.abspath("//cut_at_markers_log.txt")
file_handler = logging.FileHandler(log_path)
file_handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

def cut_strips_at_markers():
    scene = bpy.context.scene
    markers = scene.timeline_markers
    seq = scene.sequence_editor

    if not seq:
        logger.error("No sequence editor found. Aborting.")
        return

    # Sort markers by frame
    marker_frames = sorted([marker.frame for marker in markers])
    logger.info(f"Found {len(marker_frames)} markers for cutting: {marker_frames}")

    for frame in marker_frames:
        bpy.context.scene.frame_current = frame  # Set current frame to marker
        logger.info(f"Cutting at frame {frame}")

        # Select strips intersecting this frame (i.e., the frame is within their bounds)
        selected_strips = []
        for strip in seq.sequences_all:
            if strip.frame_start < frame < strip.frame_final_end:
                strip.select = True
                selected_strips.append(strip.name)
            else:
                strip.select = False

        logger.debug(f"Selected {len(selected_strips)} strips at frame {frame}: {selected_strips}")

        # Perform the cut at current frame (both sides of the blade)
        bpy.ops.sequencer.split(frame=frame, type='SOFT', side='BOTH')

    logger.info(f"✅ Completed cutting at {len(marker_frames)} markers.")

cut_strips_at_markers()
