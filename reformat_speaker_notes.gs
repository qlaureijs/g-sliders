function onOpen(e) {
  SlidesApp.getUi()
    .createAddonMenu()
    .addItem('Format Speaker Notes', 'formatSpeakerNotes')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function formatSpeakerNotes() {
  const slides = SlidesApp.getActivePresentation().getSlides();

  for (let i = 0; i < slides.length; i++) {
    const notesPage = slides[i].getNotesPage();
    const shape = notesPage.getSpeakerNotesShape();

    if (shape) {
      const textRange = shape.getText();
      const plainText = textRange.asString();

      if (plainText.trim().length > 0) {
        const style = textRange.getTextStyle();
        style.setFontSize(14);
        style.setFontFamily("Arial");
        style.setForegroundColor("#000000");
        style.setBold(false);
        style.setItalic(false);
      }
    }
  }

  SlidesApp.getUi().alert('Speaker notes formatted!');
}
