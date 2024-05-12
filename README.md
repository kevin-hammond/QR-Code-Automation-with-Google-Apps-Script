# Event QR Code Generator

This Google Apps Script generates a QR code for an event and emails it to the participant. The QR code contains a link to a form with event details.

## Setup

1. Create a new Google Apps Script project.
2. Replace `https://baseurl.com/form?` in `generateAndSendQRCode` function with the base URL of your form.
3. Replace `1fOpJriiTydFkKhRkFBMmwps09YuDBraeObXDuNBEWCg` in `generateAndSendQRCode` and `replacePlaceholder` functions with the ID of your Google Slides template.

## Usage

The script is triggered by a GET request to the web app with the following parameters:

- `eventId`: The ID of the event.
- `eventLocation`: The location of the event.
- `eventResource`: The resource for the event.
- `eventStartTime`: The start time of the event.
- `participantName`: The name of the participant.
- `email`: The email address of the person who is generating the flyer.

Example:

```
https://script.google.com/macros/s/<web-app-id>/exec?eventId=12345&eventLocation=Mars&eventResource=Lab&eventStartTime=16:00:00&participantName=John%20Doe&email=john.doe@example.com
```

## Functions

- `doGet(request)`: Handles the GET request and generates the QR code if all parameters are present.
- `generateAndSendQRCode(eventId, eventLocation, eventResource, eventStartTime, participantName, email)`: Generates the QR code and sends it to the participant.
- `convertTime(time)`: Converts the event start time to a 12-hour format.
- `replacePlaceholder(slide, placeholderText, value, isImage)`: Replaces placeholders in the Google Slides template with the event details.
- `testFunction()`: A test function to simulate a GET request.

## Testing

To test the script, run the `testFunction()` function. This function creates a mock request and logs the response.

# Label Generator Automation

This script automates the creation of a Google Slide from a provided document, replacing placeholders with specified values and emailing the generated slide as a PNG image.

## Setup

1. Replace `baseurl.com` in `doGet` function with your Firestore base URL.
2. Replace `1qq8__NpfrbPFn7Zm3YJef28JMe2wLg_mQ51jFtFSVTU` in `createSlidesFromDocument` function with your Google Slides template ID.

## API Endpoint

`https://script.google.com/macros/s/<your-script-id>/exec`

## Request Parameters

- `action`: Required. Set to `createSlide`.
- `documentId`: Required. The ID of the Firestore document.
- `assetId`: Required. The ID of the asset.
- `email`: Required. The email address to send the generated slide.

## Usage

Send a GET request with the required parameters to the API endpoint.

## Functions

- `doGet(request)`: The main function handling the GET request and calling the appropriate function based on the request parameters.
- `createSlidesFromDocument(documentId, assetId, email)`: Creates a new slide from the Google Slides template, replacing placeholders with the provided values and sending the generated slide as a PNG image to the specified email address.
- `replacePlaceholder(slide, placeholder, value, isImage)`: Replaces placeholders in the slide with the provided value. If `isImage` is true, an image is inserted; otherwise, the text is replaced.
