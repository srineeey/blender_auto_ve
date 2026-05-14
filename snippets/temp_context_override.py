import bpy

area_type = 'SEQUENCE_EDITOR' # change this to use the correct Area Type context you want to process in
areas  = [area for area in bpy.context.window.screen.areas if area.type == area_type]

if len(areas) <= 0:
    raise Exception(f"Make sure an Area of type {area_type} is open or visible in your screen!")

with bpy.context.temp_override(
    window=bpy.context.window,
    area=areas[0],
    region=[region for region in areas[0].regions if region.type == 'WINDOW'][0],
    screen=bpy.context.window.screen
):