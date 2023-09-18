# Forwarded Emails Extractor

A Python script for extracting and saving forwarded email addresses from Gmail accounts into an Excel file.

## Table of Contents
- [Introduction](#introduction)
- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
- [Usage](#usage)
- [Output](#output)
- [License](#license)

## Introduction

This Python script extracts forwarded email addresses from Gmail accounts. It uses the Gmail API to retrieve emails with subjects containing "FW" or "Fwd," extracts email addresses from the "From" lines of these emails, and saves them to an Excel file.

## Prerequisites

Before running the script, make sure you have the following prerequisites installed and set up:

- Python 3.x
- Google API Client Library (`google-auth`, `google-auth-oauthlib`, `google-auth-httplib2`)
- Google OAuth 2.0 credentials file
- Gmail API enabled for your Google account

## Getting Started

1. Clone this repository to your local machine:

```
git clone https://github.com/yourusername/forwarded-emails-extractor.git
```

2. Install the required Python packages:

pip install -r requirements.txt

3. Set up Google OAuth 2.0 credentials:
- Follow Google's guide on setting up OAuth 2.0 credentials for the Gmail API: [Getting Started with Gmail API](https://developers.google.com/gmail/api/quickstart)
- Download the `credentials.json` file and save it in the project folder.

## Usage

Run the script as follows:

```
python forwarded_emails_extractor.py
```

The script will authenticate with your Google account, retrieve forwarded emails, and save them to an Excel file.

## Output

The extracted forwarded email addresses will be saved in an Excel file named `forwarded_emails.xlsx`. Each row in the Excel file represents an extracted email address.
