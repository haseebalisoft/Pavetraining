import Image from "next/image";

import styles from "./BrandLogo.module.css";

interface BrandLogoProps {
  variant?: "full" | "compact" | "mark";
  priority?: boolean;
  className?: string;
}

export function BrandLogo({
  variant = "full",
  priority = false,
  className,
}: BrandLogoProps) {
  const size =
    variant === "full"
      ? { width: 220, height: 220 }
      : variant === "compact"
        ? { width: 148, height: 148 }
        : { width: 40, height: 40 };

  return (
    <Image
      className={`${styles.logo} ${styles[variant]} ${className ?? ""}`}
      src="/brand/pave-logo.png"
      alt="PAVE Training"
      width={size.width}
      height={size.height}
      priority={priority}
    />
  );
}
