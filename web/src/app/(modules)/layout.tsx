import Link from "next/link";

/**
 * Wraps every module page (leave, expense) with a thin bar back to the home
 * screen. The module pages themselves are unchanged apart from their API
 * prefix — they still own their whole <main>, tabs and identity strip.
 */
export default function ModulesLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <>
      <nav className="sticky top-0 z-20 border-b border-line bg-page/90 px-6 backdrop-blur sm:px-8">
        <div className="mx-auto flex w-full max-w-4xl items-center py-2">
          <Link href="/" className="text-sm font-semibold text-accent hover:underline">
            ← All requests
          </Link>
        </div>
      </nav>
      {children}
    </>
  );
}
