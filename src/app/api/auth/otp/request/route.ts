import { NextResponse } from "next/server";

import { requestEmailOtp } from "@/lib/services/otpAuthService";
import {
  PUBLIC_ERROR_MESSAGES,
  toPublicErrorPayload,
  ValidationError,
} from "@/lib/services/errorHandler";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    const body = (await request.json()) as { email?: unknown };
    const email = typeof body.email === "string" ? body.email : "";
    const result = await requestEmailOtp({ email });
    return NextResponse.json(result);
  } catch (error) {
    if (error instanceof ValidationError) {
      return NextResponse.json(
        { error: error.message, ok: false },
        { status: 400 },
      );
    }
    console.error("[otp/request]", error);
    return NextResponse.json(toPublicErrorPayload(error), { status: 500 });
  }
}
