function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const { contact, imageLink, waMsg } = data;

  const URL = here should be link.

  const whatSend = {
    username: "-usernamec",
    password: "password",
    receiverMobileNo: contact,
    message: [waMsg],
    filePathUrl: [imageLink]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(whatSend)
  };

  try {
    const response = UrlFetchApp.fetch(URL, options);
    const code = response.getResponseCode();
    const body = response.getContentText();
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.TEXT);
  } catch (e) {
    return ContentService.createTextOutput("ERROR: " + e.toString());
  }
}
