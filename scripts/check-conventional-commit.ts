const ALLOWED_TYPES = new Set([
  "feat",
  "fix",
  "refactor",
  "docs",
  "test",
  "chore",
  "ci",
  "build",
  "perf",
  "revert",
]);

const FORBIDDEN_DESCRIPTIONS = new Set([
  "cleanup",
  "misc",
  "stuff",
  "tmp",
  "updates",
  "wip",
]);

function fail(message: string): never {
  console.error(`MarkOut commit subject check failed: ${message}`);
  process.exit(1);
}

function readSubject(): string {
  const subject =
    process.env.MARKOUT_COMMIT_SUBJECT ?? process.argv.slice(2).join(" ");

  if (subject.trim().length === 0) {
    fail(
      "Provide a subject via MARKOUT_COMMIT_SUBJECT or as a command-line argument."
    );
  }

  return subject.trim();
}

function validateSubject(subject: string): void {
  const match =
    /^(?<type>[a-z]+)\((?<scope>[a-z][a-z0-9-]*)\): (?<description>.+)$/u.exec(
      subject
    );

  const type = match?.groups?.type;
  const scope = match?.groups?.scope;
  const description = match?.groups?.description;

  if (type === undefined || scope === undefined || description === undefined) {
    fail(
      "Expected `<type>(<scope>): <description>` with a non-empty scope and description."
    );
  }

  if (!ALLOWED_TYPES.has(type)) {
    fail(
      `Unsupported type \`${type}\`. Allowed types: ${Array.from(ALLOWED_TYPES).join(", ")}.`
    );
  }

  if (scope.trim().length === 0) {
    fail("The scope must be present and non-empty.");
  }

  if (description.endsWith(".")) {
    fail("The description must not end with a period.");
  }

  if (FORBIDDEN_DESCRIPTIONS.has(description.toLowerCase())) {
    fail(
      `The description \`${description}\` is too vague for this repository.`
    );
  }

  const firstCharacter = description[0];
  if (
    firstCharacter !== undefined &&
    /[A-Z]/.test(firstCharacter) &&
    firstCharacter === firstCharacter.toUpperCase()
  ) {
    fail("The description must stay lowercase after the colon.");
  }
}

const subject = readSubject();
validateSubject(subject);
console.log("MarkOut commit subject check passed.");
