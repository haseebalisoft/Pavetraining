import NextAuth from "next-auth";
import MicrosoftEntraID from "next-auth/providers/microsoft-entra-id";

export const { handlers, auth, signIn, signOut } = NextAuth({
  providers: [
    MicrosoftEntraID({
      clientId: process.env.AUTH_MICROSOFT_ENTRA_ID_ID!,
      clientSecret: process.env.AUTH_MICROSOFT_ENTRA_ID_SECRET!,
      issuer: process.env.AUTH_MICROSOFT_ENTRA_ID_ISSUER,
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
    },
  },
  callbacks: {
    async jwt({ token, profile }) {
      const profileEmail =
        typeof profile?.email === "string"
          ? profile.email
          : typeof profile?.preferred_username === "string"
            ? profile.preferred_username
            : undefined;

      if (profileEmail) {
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
