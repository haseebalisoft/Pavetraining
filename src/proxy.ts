import { NextResponse } from "next/server";

import { auth } from "@/auth";

const publicPaths = new Set([
  "/login",
  "/access-denied",
  "/preview/admin",
  "/~offline",
]);

export const proxy = auth((req) => {
  const { pathname } = req.nextUrl;
  const isLoggedIn = !!req.auth;

  // PWA service worker + workbox runtime assets must stay public.
  if (
    pathname === "/sw.js" ||
    pathname.startsWith("/swe-worker") ||
    pathname === "/manifest.webmanifest" ||
    pathname === "/manifest.webmanifest/"
  ) {
    return NextResponse.next();
  }

  if (pathname.startsWith("/api/auth")) {
    return NextResponse.next();
  }

  if (
    pathname.startsWith("/api/") &&
    !pathname.startsWith("/api/auth") &&
    !isLoggedIn
  ) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  if (!isLoggedIn && !publicPaths.has(pathname) && pathname !== "/") {
    const loginUrl = new URL("/login", req.nextUrl.origin);
    loginUrl.searchParams.set("callbackUrl", pathname);
    return NextResponse.redirect(loginUrl);
  }

  if (isLoggedIn && pathname === "/login") {
    return NextResponse.redirect(new URL("/", req.nextUrl.origin));
  }

  return NextResponse.next();
});

export const config = {
  matcher: [
    "/((?!_next/static|_next/image|favicon.ico|sw\\.js|swe-worker.*|icons/|.*\\.(?:svg|png|jpg|jpeg|gif|webp)$).*)",
  ],
};
