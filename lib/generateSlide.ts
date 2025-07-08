import pptxgen from "pptxgenjs";

export async function createPptx(): Promise<Buffer> {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: "CUSTOM", width: 13.33, height: 7.5 });
  pptx.layout = "CUSTOM";
  const slide = pptx.addSlide();

  const bgColor = "F8F9FB";
  const white = "FFFFFF";
  const blueText = "1A294B";
  const darkGreen = "58A65C";
  const lightGreen = "E8F1EB";

  slide.background = { fill: bgColor };

  const margin = 0.4;
  const contentWidth = 13.33 - 2 * margin;
  const contentHeight = 7.5 - 2 * margin;


  slide.addText("The Dependencies Dilemma", {
    x: margin,
    y: margin,
    w: contentWidth,
    h: 0.6,
    fontSize: 30,
    bold: true,
    color: blueText,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: margin + 0.7,
    w: contentWidth,
    h: 0.7,
    fill: { color: "FFFFFF" },
    rectRadius: 0.1,
    shadow: {
      type: "outer",
      angle: 90,
      blur: 14,
      offset: 0.2,
      opacity: 0.3,
      color: "000000",
    },
  });

  slide.addText(
    [
      {
        text: "The value of an initiative isn't just its immediate impact,",
        options: { fontSize: 16, italic: true, color: "000000" },
      },
      {
        text: " but what it unlocks ",
        options: { fontSize: 16, italic: true, color: "53A457" },
      },
      {
        text: "🔓",
        options: { fontSize: 16 },
      },
      {
        text: ".",
        options: { fontSize: 16, italic: true, color: "000000" },
      },
    ],
    {
      x: margin + 0.3,
      y: margin + 0.7,
      w: contentWidth - 0.6,
      h: 0.7,
      valign: "middle",
      shape: pptx.ShapeType.rect,
    }
  );

  const rightW = (contentWidth - 0.4) * 0.5;

  const topY = margin + 1.6;
  const gridH = contentHeight - 1.6;
  const leftW = (contentWidth - 0.4) * 0.5;

  const blockGap = 0.2;
  const blockH1 = (gridH - 2 * blockGap) * 0.4;
  const blockH2 = (gridH - 2 * blockGap) * 0.4;
  const blockH3 = (gridH - 2 * blockGap) * 0.2;

  const blockY1 = topY;
  const blockY2 = blockY1 + blockH1 + blockGap;
  const blockY3 = blockY2 + blockH2 + blockGap;

  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: blockY1,
    w: leftW,
    h: blockH1,
    fill: { color: "FFFFFF" },
    line: { color: "DDDDDD" },
    rectRadius: 0.1,
  });

  slide.addText(
    [
      { text: "Real-World Example\n\n", options: { fontSize: 14, bold: true, color: "1A294B" } },
      {
        text: "A fintech startup invested in comprehensive KYC infrastructure that enabled:\n\n",
        options: { fontSize: 10, color: "000000" },
      },
      { text: "• Launch in 4 new countries within 12 months\n", options: { fontSize: 10, color: "000000" } },
      { text: "• Add 3 regulated financial products\n", options: { fontSize: 10, color: "000000" } },
      { text: "• Partner with 2 major banks\n", options: { fontSize: 10, color: "000000" } },
      { text: "• Achieve compliance in weeks instead of months", options: { fontSize: 10, color: "000000" } },
    ],
    {
      x: margin + 0.2,
      y: blockY1 + 0.2,
      w: leftW - 0.4,
      h: blockH1 - 0.4,
      valign: "top",
    }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: blockY2,
    w: leftW,
    h: blockH2,
    fill: { color: "FFFFFF" },
    rectRadius: 0.1,
  });

  slide.addText("Dependency Mapping", {
    x: margin + 0.2,
    y: blockY2 + 0.2,
    w: leftW - 0.4,
    h: 0.3,
    fontSize: 14,
    bold: true,
    color: "1A294B",
  });

  const bulletColor = "4285F4";
  const bulletSize = 0.15;
  const itemFontSize = 12;
  const textStartX = margin + 0.2 + bulletSize + 0.15;
  const listYStart = blockY2 + 0.6;
  const lineHeight = 0.4;

  const items = ["Foundation capabilities vs. surface features", "Regulatory infrastructure unlocks market expansion", "Compliance systems enable product diversification"];

  items.forEach((text, idx) => {
    const y = listYStart + idx * lineHeight;

    slide.addShape(pptx.ShapeType.ellipse, {
      x: margin + 0.2,
      y: y + 0.1,
      w: bulletSize,
      h: bulletSize,
      fill: { color: bulletColor },
      line: { color: bulletColor },
    });

    slide.addText(text, {
      x: textStartX,
      y,
      w: leftW - (textStartX - margin),
      h: lineHeight,
      fontSize: itemFontSize,
      color: "000000",
      valign: "middle",
    });
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: blockY3,
    w: leftW,
    h: blockH3,
    fill: { color: white },
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: blockY3,
    w: leftW,
    h: blockH3,
    fill: { color: darkGreen },
    rectRadius: 0.1,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin + 0.1,
    y: blockY3,
    w: leftW - 0.1,
    h: blockH3,
    fill: { color: lightGreen },
        rectRadius: 0.1,
  });

  slide.addText(
    [
      { text: "Key Insight: ", options: { bold: true, color: darkGreen, fontSize: 12 } },
      {
        text: "Foundation investments create exponential value through what they unlock, not just their direct impact.",
        options: { color: "000000", fontSize: 12 },
      },
    ],
    {
      x: margin + 0.3,
      y: blockY3 + 0.2,
      w: leftW - 0.6,
      h: blockH3 - 0.4,
      valign: "top",
      align: "left",
    }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin + leftW + 0.4,
    y: topY + 0.4,
    w: rightW,
    h: gridH - 1,
    fill: { color: white },
    rectRadius: 0.1,
  });
  

  const treeX = margin + leftW + 0.4;
  const treeY = topY + 0.5;
  const treeW = rightW;

  const colCount = 4;
  const boxW = 1.2;
  const boxH = 0.4;
  const spacingX = (treeW - colCount * boxW) / (colCount + 1);
  const spacingY = 0.35;

  const levelsY = Array.from({ length: 4 }, (_, i) => treeY + 0.5 + i * (boxH + spacingY));
  const colsX = Array.from({ length: colCount }, (_, i) => treeX + spacingX + i * (boxW + spacingX));

  const colors = {
    revenue: "A52A2A",
    product: "F4B400",
    compliance: "0F9D58",
    foundation: "4285F4",
  };

  const revenue = ["Banking-as-a-Service", "White-Label Solutions", "Cross-Border Payments", "Institutional Trading"];
  const products = ["International Markets", "Business Banking", "Investment Platform", "Lending Products"];
  const compliance = ["AML Monitoring", "Regulatory Reporting", "Risk Assessment"];
  const foundation = ["KYC/Identity\nVerification"];

  function drawBox(text: string, x: number, y: number, color: string, w = boxW) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w,
      h: boxH,
      fill: { color },
    });
    slide.addText(text, {
      x,
      y,
      w,
      h: boxH,
      align: "center",
      valign: "middle",
      fontSize: 8,
      bold: true,
      color: "FFFFFF",
    });
  }

  slide.addText("Feature Enablement Tree", {
    x: treeX,
    y: treeY,
    w: treeW,
    h: 0.4,
    align: "center",
    valign: "middle",
    color: "1A294B",
    fontSize: 14,
  });

  revenue.forEach((text, i) => drawBox(text, colsX[i], levelsY[0], colors.revenue));

  products.forEach((text, i) => drawBox(text, colsX[i], levelsY[1], colors.product));

  for (let i = 0; i < colCount; i++) {
    const centerX = colsX[i] + boxW / 2;
    const fromY = levelsY[0] + boxH;
    const toY = levelsY[1];

    slide.addShape(pptx.ShapeType.line, {
      x: centerX,
      y: fromY,
      w: 0,
      h: toY - fromY,
      line: { color: colors.product, width: 1.5 },
    });
  }

  const productsLeftX = colsX[0];
  const productsRightX = colsX[3] + boxW;
  const complianceAreaCenter = (productsLeftX + productsRightX) / 2;
  const complianceTotalWidth = 3 * boxW + 2 * spacingX;
  const complianceStartX = complianceAreaCenter - complianceTotalWidth / 2;
  const complianceXs = [0, 1, 2].map((i) => complianceStartX + i * (boxW + spacingX));

  compliance.forEach((text, i) => drawBox(text, complianceXs[i], levelsY[2], colors.compliance));

  [
    { from: 0, to: 0 },
    { from: 1, to: 0 },
    { from: 2, to: 1 },
    { from: 3, to: 2 },
  ].forEach(({ from, to }) => {
    const x1 = colsX[from] + boxW / 2;
    const y1 = levelsY[1] + boxH;
    const x2 = complianceXs[to] + boxW / 2;
    const y2 = levelsY[2];

    slide.addShape(pptx.ShapeType.line, {
      x: x1,
      y: y1,
      w: x2 - x1,
      h: y2 - y1,
      line: { color: colors.compliance, width: 1.5 },
    });
  });

  const foundationW = boxW * 2 + spacingX;
  const foundationCenterX = complianceXs[1] + boxW / 2;
  const foundationX = foundationCenterX - foundationW / 2;

  drawBox(foundation[0], foundationX, levelsY[3], colors.foundation, foundationW);

  complianceXs.forEach((x) => {
    const fromX = x + boxW / 2;
    const fromY = levelsY[2] + boxH;
    const toY = levelsY[3];
    const delta = fromX - foundationCenterX;
    const toX = foundationCenterX + delta * 0.8;

    slide.addShape(pptx.ShapeType.line, {
      x: fromX,
      y: fromY,
      w: toX - fromX,
      h: toY - fromY,
      line: { color: colors.foundation, width: 1.5 },
    });
  });

  const legendItems = [
    { label: "Foundation", color: colors.foundation },
    { label: "Compliance", color: colors.compliance },
    { label: "Products", color: colors.product },
    { label: "Revenue", color: colors.revenue },
  ];

  const squareSize = 0.2;
  const labelW = 1.2;
  const itemW = squareSize + labelW;
  const totalLegendW = legendItems.length * itemW;

  const whiteBoxX = margin + leftW + 0.4;
  const whiteBoxW = rightW;
  const legendX = whiteBoxX + (whiteBoxW - totalLegendW) / 2;

  const legendY = levelsY[3] + boxH + 0.3;

  legendItems.forEach((item, i) => {
    const x = legendX + i * itemW;

    slide.addShape(pptx.ShapeType.rect, {
      x: x + squareSize,
      y: legendY,
      w: squareSize,
      h: squareSize,
      fill: { color: item.color },
    });

    slide.addText(item.label, {
      x: x + squareSize + squareSize,
      y: legendY,
      w: labelW,
      h: squareSize,
      fontSize: 10,
      valign: "middle",
      color: "000000",
    });
  });

  const result = await pptx.write({ outputType: "nodebuffer" });

  return Buffer.from(result as ArrayBuffer);
}
