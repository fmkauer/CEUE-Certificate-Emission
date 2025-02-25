# Certificate Generation System

This guide will help you generate certificates automatically using a template and a list of participants. Even if you don't have coding experience, you can follow these steps to create certificates for your event or course.

## Prerequisites

Before starting, you'll need:

1. A Google account to access Google Workspace (formerly G Suite)
2. Python installed on your computer (version 3.8 or higher)

## Step 1: Preparing Your Documents in Google Workspace

### Certificate Template (Google Docs)
1. Open Google Docs
2. Create your certificate template
3. Add placeholders for dynamic content using curly braces, for example:
   - {name} - for participant's name
   - {course} - for course name
   - {date} - for certificate date
4. Download the template as a Word document (.docx)
5. Place the downloaded file in the `input` folder and name it `Declaração de Créditos Complementares - Template.docx`

### Participant Data (Google Sheets)
1. Open Google Sheets
2. Create a spreadsheet with the following columns:
   - name: Participant's full name
   - course: Course or event name
   - date: Certificate date
3. Fill in the information for all participants
4. Download the spreadsheet as a CSV file
5. Place the downloaded file in the `data` folder and name it `Certificado CEUE (Respostas) - Form Responses 1.csv`

## Step 2: Setting Up the Project

1. Download this project to your computer
2. Open your computer's terminal or command prompt
3. Navigate to the project folder
4. Install the required software by running these commands:
   ```bash
   pip install poetry
   poetry install
   ```

## Step 3: Generating Certificates

1. Make sure your files are in the correct locations:
   - Certificate template (.docx) in the `input` folder
   - Participant data (.csv) in the `data` folder

2. Run the certificate generation program:
   ```bash
   poetry run python converter.py
   ```

3. Wait for the process to complete. You'll see progress messages in the terminal.

4. Find your generated certificates in the `output` folder. Each certificate will be named using the participant's name.

## Folder Structure

- `input/` - Place your certificate template here
- `data/` - Place your participant data CSV file here
- `output/` - Generated certificates will appear here
- `converter.py` - The main program file (don't modify unless you know Python)

## Common Issues and Solutions

1. **File Not Found Error**
   - Check if your files are in the correct folders
   - Verify the file names match exactly: `Declaração de Créditos Complementares - Template.docx` and `Certificado CEUE (Respostas) - Form Responses 1.csv`

2. **Wrong File Format**
   - Make sure the template is saved as .docx (Word document)
   - Make sure the participant data is saved as .csv (Comma Separated Values)

3. **Incorrect Placeholders**
   - Verify that the placeholders in your template match exactly: {name}, {course}, {date}
   - Check for typos and correct case sensitivity

## Need Help?

If you encounter any issues:
1. Double-check all the steps above
2. Make sure all files are in their correct locations
3. Verify file names and formats
4. Contact technical support if problems persist

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
