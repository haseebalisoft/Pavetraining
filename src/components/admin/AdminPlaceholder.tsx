import { ComingNextNotice } from "@/components/ui/States";

interface AdminPlaceholderProps {
  title: string;
  description: string;
}

export function AdminPlaceholder({ title, description }: AdminPlaceholderProps) {
  return <ComingNextNotice title={title} description={description} />;
}
