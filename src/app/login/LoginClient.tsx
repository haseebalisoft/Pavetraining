"use client";

import { useEffect, useRef, useState, type FormEvent, type ReactNode } from "react";
import { signIn } from "next-auth/react";
import { useSearchParams } from "next/navigation";

import styles from "./login.module.css";

type OtpStep = "email" | "code";

export function LoginClient({
  microsoftButton,
}: {
  microsoftButton: ReactNode;
}) {
  const searchParams = useSearchParams();
  const autoSignInTried = useRef(false);

  const [step, setStep] = useState<OtpStep>("email");
  const [email, setEmail] = useState("");
  const [code, setCode] = useState("");
  const [challenge, setChallenge] = useState<string | null>(null);
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [sending, setSending] = useState(false);
  const [verifying, setVerifying] = useState(false);

  useEffect(() => {
    const emailParam = searchParams.get("email")?.trim() || "";
    const codeParam = searchParams.get("code")?.trim() || "";
    const challengeParam = searchParams.get("challenge")?.trim() || "";

    if (emailParam) {
      setEmail(emailParam);
    }
    if (codeParam) {
      setCode(codeParam.replace(/\D/g, "").slice(0, 6));
    }
    if (challengeParam) {
      setChallenge(challengeParam);
    }
    if (emailParam && (codeParam || challengeParam)) {
      setStep("code");
      setMessage("Code loaded from your email link. Sign in to continue.");
    }
  }, [searchParams]);

  useEffect(() => {
    if (autoSignInTried.current) return;
    const emailParam = searchParams.get("email")?.trim() || "";
    const codeParam = (searchParams.get("code")?.trim() || "").replace(/\D/g, "");
    const challengeParam = searchParams.get("challenge")?.trim() || "";
    if (
      !emailParam ||
      !/^\d{6}$/.test(codeParam) ||
      !challengeParam
    ) {
      return;
    }

    autoSignInTried.current = true;
    setVerifying(true);
    setError(null);
    void (async () => {
      try {
        const result = await signIn("email-otp", {
          email: emailParam.toLowerCase(),
          code: codeParam,
          challenge: challengeParam,
          redirect: false,
        });
        if (result?.error) {
          setStep("code");
          setEmail(emailParam);
          setCode(codeParam);
          setChallenge(challengeParam);
          setError(
            "This email link is invalid or expired. Request a new code below.",
          );
          return;
        }
        let destination = "/";
        try {
          const meRes = await fetch("/api/me");
          if (meRes.ok) {
            const me = (await meRes.json()) as { redirectTo?: string };
            if (me.redirectTo === "/admin" || me.redirectTo === "/customer") {
              destination = me.redirectTo;
            }
          }
        } catch {
          // Fall back to home router.
        }
        window.location.href = destination;
      } catch (err) {
        setStep("code");
        setError(err instanceof Error ? err.message : "Sign-in failed.");
      } finally {
        setVerifying(false);
      }
    })();
  }, [searchParams]);

  async function onSendCode(event: FormEvent) {
    event.preventDefault();
    setError(null);
    setMessage(null);
    setSending(true);
    try {
      const response = await fetch("/api/auth/otp/request", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email }),
      });
      const data = (await response.json()) as {
        ok?: boolean;
        challenge?: string | null;
        message?: string;
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Could not send code.");
      }
      setChallenge(data.challenge ?? "pending");
      setCode("");
      setStep("code");
      setMessage(
        data.message ||
          "If this email is registered, a one-time code has been sent.",
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not send code.");
    } finally {
      setSending(false);
    }
  }

  async function onResendCode() {
    setError(null);
    setMessage(null);
    setSending(true);
    try {
      const response = await fetch("/api/auth/otp/request", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email }),
      });
      const data = (await response.json()) as {
        ok?: boolean;
        challenge?: string | null;
        message?: string;
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Could not send code.");
      }
      setChallenge(data.challenge ?? "pending");
      setCode("");
      setMessage(data.message || "A new code has been sent.");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not send code.");
    } finally {
      setSending(false);
    }
  }

  async function onVerifyCode(event: FormEvent) {
    event.preventDefault();
    setError(null);
    setMessage(null);
    if (!challenge || challenge === "pending") {
      setError(
        "No code was issued for this email. Check Permissions, or try another email.",
      );
      return;
    }
    setVerifying(true);
    try {
      const result = await signIn("email-otp", {
        email: email.trim().toLowerCase(),
        code: code.trim(),
        challenge,
        redirect: false,
      });
      if (result?.error) {
        setError(
          "Invalid or expired code, or this email has no portal access. Check Permissions and try again.",
        );
        return;
      }
      let destination = "/";
      try {
        const meRes = await fetch("/api/me");
        if (meRes.ok) {
          const me = (await meRes.json()) as { redirectTo?: string };
          if (me.redirectTo === "/admin" || me.redirectTo === "/customer") {
            destination = me.redirectTo;
          }
        }
      } catch {
        // Fall back to home router.
      }
      window.location.href = destination;
    } catch (err) {
      setError(err instanceof Error ? err.message : "Sign-in failed.");
    } finally {
      setVerifying(false);
    }
  }

  function goBackToEmail() {
    setStep("email");
    setCode("");
    setChallenge(null);
    setError(null);
    setMessage(null);
  }

  if (step === "code") {
    return (
      <div className={styles.otpBlock}>
        <p className={styles.otpTitle}>Enter your code</p>
        <p className={styles.otpHint}>
          We sent a 6-digit code to <strong>{email}</strong>. Check your inbox
          (and spam), or use the link in that email.
        </p>

        <form className={styles.otpForm} onSubmit={onVerifyCode}>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>6-digit code</span>
            <input
              className={styles.input}
              type="text"
              inputMode="numeric"
              pattern="[0-9]{6}"
              maxLength={6}
              autoComplete="one-time-code"
              autoFocus
              required
              value={code}
              onChange={(event) =>
                setCode(event.target.value.replace(/\D/g, "").slice(0, 6))
              }
              placeholder="123456"
            />
          </label>
          <button className={styles.button} type="submit" disabled={verifying}>
            {verifying ? "Signing in…" : "Sign in with code"}
          </button>
        </form>

        <div className={styles.otpActions}>
          <button
            type="button"
            className={styles.linkButton}
            onClick={goBackToEmail}
          >
            Use a different email
          </button>
          <button
            type="button"
            className={styles.linkButton}
            onClick={() => void onResendCode()}
            disabled={sending}
          >
            {sending ? "Sending…" : "Resend code"}
          </button>
        </div>

        {message ? <p className={styles.success}>{message}</p> : null}
        {error ? <p className={styles.formError}>{error}</p> : null}
      </div>
    );
  }

  return (
    <>
      {microsoftButton}
      <div className={styles.otpBlock}>
        <div className={styles.divider}>
          <span>or</span>
        </div>

        <p className={styles.otpTitle}>Email one-time code</p>
        <p className={styles.otpHint}>
          For customers without Microsoft. Use the same email as Permissions.
        </p>

        <form className={styles.otpForm} onSubmit={onSendCode}>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Email</span>
            <input
              className={styles.input}
              type="email"
              autoComplete="email"
              required
              value={email}
              onChange={(event) => setEmail(event.target.value)}
              placeholder="you@company.com"
            />
          </label>
          <button
            className={styles.secondaryButton}
            type="submit"
            disabled={sending}
          >
            {sending ? "Sending…" : "Send code"}
          </button>
        </form>

        {message ? <p className={styles.success}>{message}</p> : null}
        {error ? <p className={styles.formError}>{error}</p> : null}
      </div>
    </>
  );
}
