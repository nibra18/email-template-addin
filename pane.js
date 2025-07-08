function insertTemplate(fileName) {
  const baseUrl = "https://hkag.sharepoint.com/Email%20Templates";
  const fullUrl = `${baseUrl}/${fileName}`;

  fetch(fullUrl)
    .then((res) => res.text())
    .then((html) => {
      Office.context.mailbox.item.body.setSelectedDataAsync(
        html,
        { coercionType: Office.CoercionType.Html }
      );
    });
}
