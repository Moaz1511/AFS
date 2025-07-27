// Main.js
Office.onReady(() => {
  document.getElementById('convert-button').onclick = convertLaTeXInDocument;
});

async function convertLaTeXInDocument() {
  const latexRegex = /\$(.*?)\$/g;
  
  // Get all text content from presentation
  const slides = await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("shapes/text/text");
    await context.sync();
    
    // Process each slide
    slides.items.forEach(slide => {
      slide.shapes.items.forEach(shape => {
        if (shape.textFrame) {
          const text = shape.textFrame.text;
          const matches = text.match(latexRegex);
          
          if (matches) {
            matches.forEach(latexEq => {
              const cleanLatex = latexEq.replace(/\$/g, '');
              const mathML = convertToMathML(cleanLatex);
              replaceWithEquation(shape, latexEq, mathML);
            });
          }
        }
      });
    });
  });
}

function convertToMathML(latex) {
  // Use a library like latex2mathml or call a service
  return latex2mathml(latex);
}

async function replaceWithEquation(shape, originalText, mathML) {
  // Office JS API to insert equation
  await Office.context.document.setSelectedDataAsync(mathML, {
    coercionType: Office.CoercionType.Ooxml
  });
}

Office.context.document.addHandlerAsync(
  Office.EventType.DocumentSelectionChanged,
  handleSelectionChange
);

function handleSelectionChange(eventArgs) {
  const selection = Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text,
    (result) => {
      if (result.value.match(/\$(.*?)\$/)) {
        showConversionUI();
      }
    }
  );
}

function applyEquationStyle(equationOoxml) {
  // Modify OOXML to include styling
  return equationOoxml.replace(
    '</m:oMath>',
    '<m:ctrlPr><m:sty m:val="p"/></m:ctrlPr></m:oMath>'
  );
}