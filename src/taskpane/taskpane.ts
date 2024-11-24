/* global PowerPoint console */

export async function insertText(text: string) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const textBox = slide.shapes.addTextBox(text);
      textBox.fill.setSolidColor("white");
      textBox.lineFormat.color = "black";
      textBox.lineFormat.weight = 1;
      textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      console.log(context.application.context.debugInfo);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function extractSlideAsImage() {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      await context.sync();
      console.log(image.value);
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
