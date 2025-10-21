# Daycare Supporter(Daycare Docs creater)

A Windows desktop app to generate polished childcare notices in minutes.
Enter event details in a single window and get TXT and styled DOCX output with seasonal greetings and event-specific intros that rotate automatically.
This project was originally developed to support my family’s daycare in Korea and is currently available in Korean only.

- UI: Tkinter
- Document engine: python-docx (fonts/margins/rules/bullets/footnotes)
- Packaging: PyInstaller (single EXE)

“Consistent tone and layout, produced in under a minute.”

## Features

- Event types: Field Trip / Class Observation / Picnic / Health/Immunization / General Notice
- Seasonal greeting rotation based on date (Spring/Summer/Autumn/Winter)
- Intro rotation per event type (3 variants each)


## Screenshots
- UI form
<img width="761" height="753" alt="Image" src="https://github.com/user-attachments/assets/f425feb1-ebc2-491c-ab5a-fe05b1797814" />

- DOCX output
<img width="447" height="669" alt="Image" src="https://github.com/user-attachments/assets/a6ca0fd0-a2c1-4bf8-a5d7-bbed8f4e251e" />

- Banner example(optional)
<img width="1536" height="1024" alt="Image" src="https://github.com/user-attachments/assets/dfd62239-126f-4c31-9eb5-a7cd33047363" />

## Tech Stack

- Python 3.13, Tkinter, python-docx, lxml, PyInstaller

## How to use

1. Select Event Type and click Load Defaults to prefill sample values.
2. Edit date/time/location/materials as needed.
(Optional) Choose a banner image.
3. Click Generate TXT or Generate TXT + DOCX.
4. Click Open Output Folder to view results.
5. Use Copy Notice to paste plain text to chat apps.

## Development Notes

This project was developed as a solo side project.
I occasionally used generative AI (ChatGPT) as a coding assistant for boilerplate and debugging.
All feature design, architecture, and final testing were performed manually.

## License
This project is licensed under the MIT License
