# Adding Alt-Text to images in Keynote presentations

Work in progress.

AppleScript `get_images.scpt` will return a list of the images in the Keynote presentation currently open in Keynote.

Proposed workflow:

- Export Keynote presentation to Powerpoint.
- Run `auto_alt_text.py` to generate text file with image descriptions.
- Run an AppleScript (not yet written) that uses the generated text file to add the descriptions to the images in the Keynote presentation.
