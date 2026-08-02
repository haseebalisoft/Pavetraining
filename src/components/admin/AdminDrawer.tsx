"use client";

import type { ReactNode } from "react";

import { SlideOverPanel } from "@/components/ui/SlideOverPanel";

interface AdminDrawerProps {
  open: boolean;
  title: string;
  onClose: () => void;
  children: ReactNode;
  footer: ReactNode;
  wide?: boolean;
}

export function AdminDrawer({
  open,
  title,
  onClose,
  children,
  footer,
  wide = false,
}: AdminDrawerProps) {
  return (
    <SlideOverPanel
      open={open}
      title={title}
      onClose={onClose}
      footer={footer}
      wide={wide}
    >
      {children}
    </SlideOverPanel>
  );
}
