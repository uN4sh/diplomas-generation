# Automatic diplomas generation and sending

Generates named diplomas from a template and a GSheets table then sends it to the students. Requires GSheet and GMail API — or similar.
> 11.03.2021 @ France — @vdElyn

- Get your Sheets API Credential JSON file [from here](https://gspread.readthedocs.io/en/latest/)
- Get your Mail API Credential JSON file and enable the Gmail API [from here](https://developers.google.com/gmail/api/quickstart/python)
- Complete `config.yaml` file with your Sheets and Gmail data and specifications
- Edit `message_text.html` to edit your mail message
- Run the python `generate_diplomas.py` file

### Required modules
- `pyyaml` to parse .yaml files
- `PIL` for the image processing
- `Google Sheets API` and `gspread` to deal with the Sheet
- `Google Mail API` to send e-mails to students


```python
$ pip install gspread pyyaml Pillow
$ pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```