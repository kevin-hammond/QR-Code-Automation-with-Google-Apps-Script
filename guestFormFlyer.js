// Template: https://docs.google.com/presentation/d/1fOpJriiTydFkKhRkFBMmwps09YuDBraeObXDuNBEWCg/edit?usp=sharing

function doGet(request) {
  var eventId = request.parameter.eventId;
  var eventLocation = request.parameter.eventLocation;
  var eventResource = request.parameter.eventResource;
  var eventStartTime = request.parameter.eventStartTime;
  var participantName = request.parameter.participantName;
  var email = request.parameter.email;

  if (
    eventId &&
    eventLocation &&
    eventResource &&
    eventStartTime &&
    participantName &&
    email
  ) {
    generateAndSendQRCode(
      eventId,
      eventLocation,
      eventResource,
      eventStartTime,
      participantName,
      email
    );
    return ContentService.createTextOutput(
      "QR code generated and sent successfully."
    );
  } else {
    return ContentService.createTextOutput(
      "Invalid Request. Check your parameters."
    );
  }
}

function generateAndSendQRCode(
  eventId,
  eventLocation,
  eventResource,
  eventStartTime,
  participantName,
  email
) {
  var flutterFlowUrl =
    "https://baseurl.com/form?" +
    "eventId=" +
    encodeURIComponent(eventId) +
    "&eventStartTime=" +
    encodeURIComponent(eventStartTime) +
    "&participantName=" +
    encodeURIComponent(participantName) +
    "&eventResource=" +
    encodeURIComponent(eventResource) +
    "&eventLocation=" +
    encodeURIComponent(eventLocation);

  var qrCodeUrl =
    "https://qrcode.tec-it.com/API/QRCode?data=" +
    encodeURIComponent(flutterFlowUrl);

  // Replace with the ID of your Google Slides template
  var templateId = "1fOpJriiTydFkKhRkFBMmwps09YuDBraeObXDuNBEWCg";

  // Get the presentation and the template slide
  var presentation = SlidesApp.openById(templateId);
  var templateSlide = presentation.getSlides()[0]; // Assuming the template slide is the first slide

  // Duplicate the template slide
  var slide = presentation.insertSlide(
    presentation.getSlides().length,
    templateSlide
  );

  // Convert the time format
  var convertedTime = convertTime(eventStartTime);

  // Replace placeholders in the slide
  replacePlaceholder(slide, "{participantName}", participantName, false);
  replacePlaceholder(slide, "{eventResource}", eventResource, false);
  replacePlaceholder(slide, "{eventStartTime}", convertedTime, false);
  replacePlaceholder(slide, "QR_CODE_IMAGE_PLACEHOLDER", qrCodeUrl, true);

  // Save and close the presentation to allow changes to propagate
  presentation.saveAndClose();

  // Wait for 3 seconds to allow changes to propagate
  Utilities.sleep(3000);

  // Re-open the presentation to fetch the thumbnail
  presentation = SlidesApp.openById(templateId);
  var presentationId = templateId;
  var slideId = slide.getObjectId(); // Get the object ID of the newly created slide
  var thumbnailSize = "LARGE"; // Set the desired thumbnail size; valid values are "SMALL", "MEDIUM", and "LARGE"
  var url =
    "https://slides.googleapis.com/v1/presentations/" +
    presentationId +
    "/pages/" +
    slideId +
    "/thumbnail?thumbnailProperties.thumbnailSize=" +
    thumbnailSize;
  var response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken(),
    },
  });
  var thumbnailUrl = JSON.parse(response.getContentText()).contentUrl;
  var imageBlob = UrlFetchApp.fetch(thumbnailUrl).getBlob().getAs("image/png");
  var emailBody =
    "Here is the QR code and link for the event:<br><br>" +
    "Event ID: " +
    eventId +
    "<br>" +
    "Event Location: " +
    eventLocation +
    "<br>" +
    "Event Resource: " +
    eventResource +
    "<br>" +
    "Event Start Time: " +
    eventStartTime +
    "<br>" +
    "Participant Name: " +
    participantName +
    "<br><br>" +
    "Link: <a href='" +
    flutterFlowUrl +
    "'>" +
    flutterFlowUrl +
    "</a>";

  // Email the image to the provided email
  MailApp.sendEmail({
    to: email,
    subject: "Event QR Code: " + participantName,
    htmlBody: emailBody,
    attachments: [imageBlob],
  });
}

function convertTime(time) {
  var parts = time.split(":");
  var hours = parseInt(parts[0]);
  var minutes = parts[1];
  var ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12;
  hours = hours ? hours : 12; // Convert 0 to 12
  return hours + ":" + minutes + " " + ampm;
}

function replacePlaceholder(slide, placeholderText, value, isImage) {
  var elements = slide.getPageElements();
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      var shape = element.asShape();
      if (shape.getText().asString().includes(placeholderText)) {
        if (isImage) {
          var updatedText = shape.getText().replaceAllText(placeholderText, "");
          shape.getText().setText(updatedText);

          var response = UrlFetchApp.fetch(value);
          var blob = response.getBlob();
          var image = slide.insertImage(blob);
          image.alignOnPage(SlidesApp.AlignmentPosition.CENTER);
        } else {
          var updatedText = shape
            .getText()
            .replaceAllText(placeholderText, value);
          shape.getText().setText(updatedText);
        }
      }
    }
  }
}

function replacePlaceholder(slide, placeholder, value, isImage) {
  var shapes = slide.getShapes();

  for (var i = 0; i < shapes.length; i++) {
    var shape = shapes[i];
    var text = shape.getText();

    if (text && text.asString().indexOf(placeholder) !== -1) {
      if (isImage) {
        // Store the size and position of the placeholder shape
        var width = shape.getWidth();
        var height = shape.getHeight();
        var left = shape.getLeft();
        var top = shape.getTop();

        // Remove the placeholder shape
        shape.remove();

        // Add the image to the slide
        var image = slide.insertImage(value);

        // Set the image size and position
        image.setWidth(width);
        image.setHeight(height);
        image.setLeft(left);
        image.setTop(top);
      } else {
        text.clear();
        text.insertText(0, value);
      }
    }
  }
}

function testFunction() {
  var mockRequest = {
    parameter: {
      eventId: "123456",
      eventLocation: "Mars Base",
      eventResource: "Laboratory 42",
      eventStartTime: "16:00:00",
      participantName: "Tyler Durden",
      email: "tyler@fightclub.com",
    },
  };

  var response = doGet(mockRequest);
  Logger.log("Response: " + response.getContent());
}
