import type { NextAuthConfig } from "next-auth";
import { NextResponse } from "next/server";

function publicOrigin(requestUrl: URL): string {
  const configured = process.env.AUTH_URL ?? process.env.NEXTAUTH_URL;
  if (configured) {
    try {
      return new URL(configured).origin;
    } catch {
      // ignore invalid env
    }
  }
  return requestUrl.origin;
}

export const authConfig: NextAuthConfig = {
  pages: {
    signIn: "/login",
  },
  session: {
    strategy: "jwt",
  },
  providers: [],
  callbacks: {
    authorized({ auth, request }) {
      const { pathname } = request.nextUrl;
      const isAuthPage =
        pathname.startsWith("/login") || pathname.startsWith("/register");
      const isProtected =
        pathname.startsWith("/dashboard") ||
        pathname.startsWith("/spreadsheet") ||
        pathname.startsWith("/api/spreadsheets");

      if (isAuthPage && auth?.user) {
        return NextResponse.redirect(
          new URL("/dashboard", publicOrigin(request.nextUrl)),
        );
      }

      if (isProtected && !auth?.user) {
        return false;
      }

      return true;
    },
    jwt({ token, user }) {
      if (user?.id) {
        token.id = user.id;
      }
      return token;
    },
    session({ session, token }) {
      if (session.user && token.id) {
        session.user.id = token.id as string;
      }
      return session;
    },
  },
  trustHost: true,
};