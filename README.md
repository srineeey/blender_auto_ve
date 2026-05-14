# blender_auto_ve

A set of scripts to leverage blenders python API for automatic video editing

## Idea

Manually editing videos can be tedious, especially for long footage where most of it is boring. This script automates the editing process by creating a video in which interesting bits are kept and boring intermediat parts are sped up - an alternating time warp so to speak. Besides manually putting markers for transitions, the editing process is fully automatic.

### Usage

#### Manual Steps

##### Open the Blender Video Editor (VE)


##### Set up a text editor panel (for running the script later)


##### Load the script inside the text editor


##### Import all video footage into the timeline

It does not matter if they have gaps inbetween - the script will glue them together later. But optionally one can right click and sleect remove gaps. 


##### Add markers

These should designate sections to keep or time warp.
The first marker (n=0) should be at the start of the footage. 
The section until the next marker (n=1) will be kept. 
The section from (n=1) to (n=2) will be sped up. 
This means that all sections n=odd to n=even will be sped up.

Sometimes the video preview is not enabled. To select a section to prerender click on the sequencer, press P, and then hoch and select the area to preview

To set a marker select the frame on the slider in the sequencer and press M. A dashed line with a small triangle on the bottom should appear. Dont worry about marker names - the script will rename them


#### Automatic Steps - handled by script

- Cut the fottage at marker locations
- speed up sections - gaps in the footage will appear and or widen!
- rearrange cut up strips to remove gaps inbetween


#### Final Step - Rendering

- Make sure that the Output options in the Scene panel are set correctly, especially the file format.
- Then hit Render - Render Animation
