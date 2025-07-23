import { NextResponse } from "next/server";
import { createPptx } from "@/lib/generateSlide";

export async function GET() {
  const pptxBuffer = await createPptx();

  return new NextResponse(pptxBuffer, {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-Disposition": "attachment; filename=Dependencies_Dilemma_Slide.pptx",
    },
  });
}
