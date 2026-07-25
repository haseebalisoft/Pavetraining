"use client";

import { AdminToastProvider } from "@/components/admin/AdminToast";
import type { ReactNode } from "react";

export function AdminProviders({ children }: { children: ReactNode }) {
  return <AdminToastProvider>{children}</AdminToastProvider>;
}
