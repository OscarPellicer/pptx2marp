# TODOs

In the test file everything.pptx, there are some issues still to be fixed:

- The background image is not being extracted.
- ~~There are many math symbols that are not being read~~ -> Fixed
- ~~Sometimes math text does not get converted to the latex command, and the actual greek symbol remains in the text. Fortunately, they do not crash python-pptx.~~ -> Fixed
- ~~The table is not being converted.~~ -> Fixed
- Shapes are not being converted.
- Look into MARP styles and create a custom style in an external css (when i tried, it was not working). Then:
  - Have a custom style for the title, subtitle, and normal text, so that title is always on top of the slide, subtitle is below the title, and normal text is below the subtitle.
  - Use a specific font that I like for all elements.
