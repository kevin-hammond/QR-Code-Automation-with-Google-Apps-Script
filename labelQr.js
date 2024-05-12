// Template: https://docs.google.com/presentation/d/1qq8__NpfrbPFn7Zm3YJef28JMe2wLg_mQ51jFtFSVTU/edit?usp=sharing

function doGet(request) {
  var action = request.parameter.action;
  var documentId = request.parameter.documentId;
  var assetId = request.parameter.assetId;
  var email = request.parameter.email; // Retrieve email from the request

  if (action === "createSlide" && documentId && assetId && email) {
    // Call the function to create the slide
    createSlidesFromDocument(documentId, assetId, email); // Pass email as an argument
    return ContentService.createTextOutput("Slide created successfully.");
  } else {
    return ContentService.createTextOutput(
      "Invalid Request. Check your parameters."
    );
  }
}

function createSlidesFromDocument(documentId, assetId, email) {
  // Replace with the ID of your Google Slides template
  var templateId = "1qq8__NpfrbPFn7Zm3YJef28JMe2wLg_mQ51jFtFSVTU";

  // Concatenate the Firestore document ID with the given URL
  var url = "https://baseurl.com/" + documentId;

  // Generate the QR code image URL
  var qrCodeUrl =
    "https://qrcode.tec-it.com/API/QRCode?data=" + encodeURIComponent(url);

  // Get the presentation and the template slide
  var presentation = SlidesApp.openById(templateId);
  var templateSlide = presentation.getSlides()[0]; // Assuming the template slide is the first slide

  // Duplicate the template slide
  var slide = presentation.insertSlide(
    presentation.getSlides().length,
    templateSlide
  );

  // Replace placeholders in the slide
  replacePlaceholder(slide, "QR_CODE_IMAGE_PLACEHOLDER", qrCodeUrl, true);
  replacePlaceholder(slide, "XX-XXX", assetId);

  // Save and close the presentation to allow changes to propagate
  presentation.saveAndClose();

  // Wait for 15 seconds to allow changes to propagate
  Utilities.sleep(15000);

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

  // Email the image to the provided email
  MailApp.sendEmail({
    to: email,
    subject: "New Label: " + assetId,
    body: "Here ya go!",
    attachments: [imageBlob],
  });
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
