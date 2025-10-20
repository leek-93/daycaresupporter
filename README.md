Daycare Supporter(Daycare Docs creater)

A Windows desktop app to generate polished childcare notices in minutes.
Enter event details in a single window and get TXT and styled DOCX output with seasonal greetings and event-specific intros that rotate automatically.
This project was originally developed to support my family’s daycare in Korea and is currently available in Korean only.

UI: Tkinter
Document engine: python-docx (fonts/margins/rules/bullets/footnotes)
Packaging: PyInstaller (single EXE)

“Consistent tone and layout, produced in under a minute.”

Features

Event types: Field Trip / Class Observation / Picnic / Health/Immunization / General Notice
Seasonal greeting rotation based on date (Spring/Summer/Autumn/Winter)
Intro rotation per event type (3 variants each)


Screenshots
UI form

DOCX output

Banner example(optional)


Tech Stack

Python 3.13, Tkinter, python-docx, lxml, PyInstaller

How to use

Select Event Type and click Load Defaults to prefill sample values.
Edit date/time/location/materials as needed.
(Optional) Choose a banner image.
Click Generate TXT or Generate TXT + DOCX.
Click Open Output Folder to view results.
Use Copy Notice to paste plain text to chat apps.

Development Notes

This project was developed as a solo side project.
I occasionally used generative AI (ChatGPT) as a coding assistant for boilerplate and debugging.
All feature design, architecture, and final testing were performed manually.

License
This project is licensed under the MIT License
