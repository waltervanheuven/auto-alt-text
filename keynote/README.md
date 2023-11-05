# Adding Alt-Text to images in Keynote presentations

Work in progress.

AppleScript `get_images.scpt` will list the images in the Keynote presentation currently open in Keynote.

Proposed workflow:

- Export Keynote presentation to Powerpoint
- Run `auto_alt_text.py` to generate text file with image descriptions
- Run an AppleScript (not yet written) that uses the generated text file to add the image descriptions to the image in the Keynote presentation.
