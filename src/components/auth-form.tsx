"use client";

import Link from "next/link";
import { signIn } from "next-auth/react";
import { FormEvent, useState } from "react";

interface AuthFormProps {
  mode: "login" | "register";
}

export function AuthForm({ mode }: AuthFormProps) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [name, setName] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setBusy(true);
    setError(null);

    try {
      if (mode === "register") {
        const response = await fetch("/api/auth/register", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ email, password, name: name || undefined }),
        });
        const payload = await response.json();
        if (!response.ok) {
          throw new Error(payload.error ?? "Registration failed");
        }
      }

      const result = await signIn("credentials", {
        email,
        password,
        redirect: false,
      });

      if (!result || result.error) {
        throw new Error(
          mode === "login" ? "Invalid email or password" : "Login after registration failed",
        );
      }

      // Full navigation ensures the session cookie is visible to middleware.
      window.location.assign("/dashboard");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Authentication failed");
    } finally {
      setBusy(false);
    }
  }

  return (
    <form
      onSubmit={handleSubmit}
      className="w-full max-w-md space-y-4 rounded-xl border border-zinc-200 bg-white p-6 text-zinc-900 shadow-sm"
    >
      <div>
        <h1 className="text-2xl font-semibold text-zinc-900">
          {mode === "login" ? "Log in" : "Create account"}
        </h1>
        <p className="text-sm text-zinc-600">
          {mode === "login"
            ? "Access your spreadsheets"
            : "Register to start editing spreadsheets locally"}
        </p>
      </div>

      {mode === "register" ? (
        <label className="block text-sm">
          <span className="mb-1 block text-zinc-700">Name (optional)</span>
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-zinc-900"
            autoComplete="name"
          />
        </label>
      ) : null}

      <label className="block text-sm" htmlFor="auth-email">
        <span className="mb-1 block text-zinc-700">Email</span>
        <input
          id="auth-email"
          type="email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          required
          className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-zinc-900"
          autoComplete="email"
        />
      </label>

      <label className="block text-sm" htmlFor="auth-password">
        <span className="mb-1 block text-zinc-700">Password</span>
        <input
          id="auth-password"
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          required
          minLength={mode === "register" ? 8 : 1}
          className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-zinc-900"
          autoComplete={
            mode === "register" ? "new-password" : "current-password"
          }
        />
      </label>

      {error ? (
        <p className="rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">
          {error}
        </p>
      ) : null}

      <button
        type="submit"
        disabled={busy}
        className="w-full rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-60"
      >
        {busy
          ? "Please wait…"
          : mode === "login"
            ? "Log in"
            : "Register"}
      </button>

      <p className="text-center text-sm text-zinc-600">
        {mode === "login" ? (
          <>
            No account?{" "}
            <Link href="/register" className="text-blue-600 hover:underline">
              Register
            </Link>
          </>
        ) : (
          <>
            Already registered?{" "}
            <Link href="/login" className="text-blue-600 hover:underline">
              Log in
            </Link>
          </>
        )}
      </p>
    </form>
  );
}