import NextAuth from "next-auth";
import Credentials from "next-auth/providers/credentials";
import MicrosoftEntraID from "next-auth/providers/microsoft-entra-id";

export const { handlers, auth, signIn, signOut } = NextAuth({
  providers: [
    MicrosoftEntraID({
      clientId: process.env.AUTH_MICROSOFT_ENTRA_ID_ID!,
      clientSecret: process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET!,
      issuer: process.env.AUTH_MICROSOFT_ENTRA_ID_ISSUER,
    }),
    Credentials({
      id: "email-otp",
      name: "Email OTP",
      credentials: {
        email: { label: "Email", type: "email" },
        code: { label: "Code", type: "text" },
        challenge: { label: "Challenge", type: "text" },
      },
      async authorize(credentials) {
        const { verifyEmailOtp } = await import(
          "@/lib/services/otpAuthService"
        );
        try {
          const verified = await verifyEmailOtp({
            email:
              typeof credentials?.email === "string"
                ? credentials.email
                : null,
            code:
              typeof credentials?.code === "string" ? credentials.code : null,
            challenge:
              typeof credentials?.challenge === "string"
                ? credentials.challenge
                : null,
          });
          return {
            id: verified.email,
            email: verified.email,
            name: verified.email.split("@")[0] ?? verified.email,
          };
        } catch {
          return null;
        }
      },
    }),
  ],
  session: {
    strategy: "jwt",
  },
  pages: {
    signIn: "/login",
    error: "/login",
  },
  events: {
    async signIn({ user }) {
      const email = user.email?.trim().toLowerCase();
      if (!email) return;
      // Never block the OAuth callback on SharePoint audit writes.
      void (async () => {
        try {
          const { getMeContext } = await import(
            "@/lib/services/customerContextService"
          );
          const { logLogin } = await import("@/lib/services/auditLogService");
          const me = await getMeContext(email);
          if (!me) {
            await logLogin({
              userEmail: email,
              success: false,
              errorMessage: "No active permission for this account.",
            });
            return;
          }
          await logLogin({
            userEmail: email,
            roleType: me.roleLabel ?? me.role,
            company: me.companyName,
            success: true,
          });
        } catch (error) {
          console.error("[audit] signIn event logging failed", error);
        }
      })();
    },
  },
  callbacks: {
    async jwt({ token, user, profile }) {
      const profileEmail =
        typeof profile?.email === "string"
          ? profile.email
          : typeof profile?.preferred_username === "string"
            ? profile.preferred_username
            : undefined;

      if (typeof user?.email === "string" && user.email.trim()) {
        token.email = user.email.toLowerCase();
      } else if (profileEmail) {
        token.email = profileEmail.toLowerCase();
      }

      return token;
    },
    async session({ session, token }) {
      if (session.user && typeof token.email === "string") {
        session.user.email = token.email;
      }
      return session;
    },
  },
  trustHost: true,
});
