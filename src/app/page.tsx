import { auth } from "@/lib/auth";
import Link from "next/link";
import { redirect } from "next/navigation";

export default async function HomePage() {
  const session = await auth();
  if (session?.user) {
    redirect("/dashboard");
  }

  return (
    <main className="flex min-h-screen flex-col items-center justify-center gap-6 bg-zinc-50 px-4 text-center">
      <div>
        <p className="text-sm font-medium uppercase tracking-wide text-blue-700">
          Local-first spreadsheet editor
        </p>
        <h1 className="mt-2 text-4xl font-semibold text-zinc-900">
          Excel-style editing in your browser
        </h1>
        <p className="mt-3 max-w-xl text-zinc-600">
          Upload `.xlsx` files, edit cells in a familiar grid, and export your
          work — all running locally with Docker and PostgreSQL.
        </p>
      </div>
      <div className="flex gap-3">
        <Link
          href="/register"
          className="rounded-lg bg-blue-600 px-5 py-2.5 text-sm font-medium text-white hover:bg-blue-700"
        >
          Get started
        </Link>
        <Link
          href="/login"
          className="rounded-lg border border-zinc-300 bg-white px-5 py-2.5 text-sm font-medium text-zinc-900 hover:bg-zinc-50"
        >
          Log in
        </Link>
      </div>
    </main>
  );
}