type Environment = Readonly<Record<string, string | undefined>>;

interface RepositoryCoordinates {
  owner: string;
  repo: string;
}

interface RulesetAuditResult {
  hasAutomationBypass: boolean;
  hasHumanBypass: boolean;
  hasReleaseProductionCondition: boolean;
  hasRequiredRuleTypes: boolean;
  hasUnexpectedBypass: boolean;
  name: string;
}

const DEFAULT_REPOSITORY = "Schoenfeld-Solutions/markout";
const GITHUB_API_BASE_URL = "https://api.github.com";
const REQUIRED_BRANCHES = ["main", "release/production"] as const;
const REQUIRED_GITHUB_PAGES_POLICIES = ["main", "release/production"] as const;
const REQUIRED_RELEASE_RULE_TYPES = ["creation", "deletion", "update"] as const;
const RELEASE_PRODUCTION_REF = "refs/heads/release/production";

async function main(): Promise<void> {
  const environment = process.env;
  const errors: string[] = [];
  const warnings: string[] = [];
  const token = readToken(environment, errors);
  const repository = readRepository(environment, errors);

  if (token.length === 0 || repository === undefined) {
    reportAuditResult(errors, warnings);
    return;
  }

  await checkRequiredBranches(token, repository, errors);
  await checkEnvironments(token, repository, errors);
  await checkGitHubPagesBranchPolicies(token, repository, errors);
  await checkReleaseProductionRulesets(token, repository, environment, errors);

  reportAuditResult(errors, warnings);
}

async function checkRequiredBranches(
  token: string,
  repository: RepositoryCoordinates,
  errors: string[]
): Promise<void> {
  for (const branchName of REQUIRED_BRANCHES) {
    await githubApiRequest(
      token,
      repository,
      `/branches/${encodeURIComponent(branchName)}`,
      `branch ${branchName}`,
      errors
    );
  }
}

async function checkEnvironments(
  token: string,
  repository: RepositoryCoordinates,
  errors: string[]
): Promise<void> {
  const environmentsPayload = await githubApiRequest(
    token,
    repository,
    "/environments",
    "repository environments",
    errors
  );

  if (environmentsPayload === undefined) {
    return;
  }

  const environments = readArrayProperty(
    environmentsPayload,
    "environments",
    "repository environments",
    errors
  );
  const environmentNames = new Set(
    environments
      .map((environment) => readOptionalStringProperty(environment, "name"))
      .filter((name): name is string => name !== undefined)
  );

  if (!environmentNames.has("github-pages")) {
    errors.push("Missing required github-pages environment.");
  }

  const productionPromotion = environments.find(
    (environment) =>
      readOptionalStringProperty(environment, "name") === "production-promotion"
  );

  if (productionPromotion === undefined) {
    errors.push("Missing required production-promotion environment.");
    return;
  }

  if (!hasRequiredReviewerProtection(productionPromotion)) {
    errors.push(
      "production-promotion must require at least one reviewer before promotion can push production."
    );
  }
}

async function checkGitHubPagesBranchPolicies(
  token: string,
  repository: RepositoryCoordinates,
  errors: string[]
): Promise<void> {
  const branchPoliciesPayload = await githubApiRequest(
    token,
    repository,
    "/environments/github-pages/deployment-branch-policies",
    "github-pages deployment branch policies",
    errors
  );

  if (branchPoliciesPayload === undefined) {
    return;
  }

  const branchPolicies = readArrayProperty(
    branchPoliciesPayload,
    "branch_policies",
    "github-pages deployment branch policies",
    errors
  );
  const policyNames = new Set(
    branchPolicies
      .map((policy) => readOptionalStringProperty(policy, "name"))
      .filter((name): name is string => name !== undefined)
  );

  for (const requiredPolicy of REQUIRED_GITHUB_PAGES_POLICIES) {
    if (!policyNames.has(requiredPolicy)) {
      errors.push(
        `github-pages deployment branch policies must include ${requiredPolicy}.`
      );
    }
  }
}

async function checkReleaseProductionRulesets(
  token: string,
  repository: RepositoryCoordinates,
  environment: Environment,
  errors: string[]
): Promise<void> {
  if ((environment.MARKOUT_REQUIRE_RELEASE_BOT ?? "true") === "false") {
    return;
  }

  const releaseBotAppId = environment.MARKOUT_RELEASE_BOT_APP_ID?.trim() ?? "";
  const releaseBotPrivateKey =
    environment.MARKOUT_RELEASE_BOT_PRIVATE_KEY?.trim() ?? "";

  if (releaseBotAppId.length === 0) {
    errors.push("MARKOUT_RELEASE_BOT_APP_ID is required.");
  }

  if (releaseBotPrivateKey.length === 0) {
    errors.push("MARKOUT_RELEASE_BOT_PRIVATE_KEY is required.");
  }

  const rulesetsPayload = await githubApiRequest(
    token,
    repository,
    "/rulesets",
    "repository rulesets",
    errors
  );

  if (rulesetsPayload === undefined) {
    return;
  }

  if (!Array.isArray(rulesetsPayload)) {
    errors.push("Repository rulesets response must be an array.");
    return;
  }

  const detailedRulesets = await Promise.all(
    rulesetsPayload.map(async (rulesetSummary) => {
      const rulesetId = readOptionalNumberProperty(rulesetSummary, "id");
      if (rulesetId === undefined) {
        return undefined;
      }

      return await githubApiRequest(
        token,
        repository,
        `/rulesets/${rulesetId}`,
        `repository ruleset ${rulesetId}`,
        errors
      );
    })
  );
  const rulesetValidationErrors = validateReleaseProductionRulesets(
    detailedRulesets.filter(
      (detailedRuleset): detailedRuleset is unknown =>
        detailedRuleset !== undefined
    ),
    releaseBotAppId
  );

  errors.push(...rulesetValidationErrors);
}

export function validateReleaseProductionRulesets(
  detailedRulesets: readonly unknown[],
  releaseBotAppId: string
): string[] {
  const errors: string[] = [];
  const matchingRulesets: RulesetAuditResult[] = [];

  for (const detailedRuleset of detailedRulesets) {
    const auditResult = auditRuleset(detailedRuleset, releaseBotAppId);
    if (
      auditResult.hasReleaseProductionCondition &&
      auditResult.hasRequiredRuleTypes
    ) {
      matchingRulesets.push(auditResult);
    }
  }

  if (matchingRulesets.length === 0) {
    errors.push(
      "release/production must have an active ruleset that blocks creation, update, and deletion."
    );
    return errors;
  }

  if (!matchingRulesets.some((ruleset) => ruleset.hasAutomationBypass)) {
    errors.push(
      "release/production ruleset must allow the markout-release-bot GitHub App as its only automation bypass."
    );
  }

  if (matchingRulesets.some((ruleset) => ruleset.hasHumanBypass)) {
    errors.push(
      "release/production ruleset must not include User, Team, or admin bypass actors."
    );
  }

  if (matchingRulesets.some((ruleset) => ruleset.hasUnexpectedBypass)) {
    errors.push(
      "release/production ruleset must not include bypass actors other than the markout-release-bot GitHub App."
    );
  }

  return errors;
}

export function auditRuleset(
  ruleset: unknown,
  expectedReleaseBotAppId: string
): RulesetAuditResult {
  const name = readOptionalStringProperty(ruleset, "name") ?? "unnamed";
  const enforcement = readOptionalStringProperty(ruleset, "enforcement");
  const target = readOptionalStringProperty(ruleset, "target");
  const rules = readOptionalArrayProperty(ruleset, "rules");
  const bypassActors = readOptionalArrayProperty(ruleset, "bypass_actors");
  const ruleTypes = new Set(
    rules
      .map((rule) => readOptionalStringProperty(rule, "type"))
      .filter((ruleType): ruleType is string => ruleType !== undefined)
  );

  return {
    hasAutomationBypass: bypassActors.some((actor) =>
      isExpectedReleaseBotBypass(actor, expectedReleaseBotAppId)
    ),
    hasHumanBypass: bypassActors.some((actor) => isHumanBypass(actor)),
    hasReleaseProductionCondition:
      enforcement === "active" &&
      target === "branch" &&
      rulesetIncludesReleaseProduction(ruleset),
    hasRequiredRuleTypes: REQUIRED_RELEASE_RULE_TYPES.every((ruleType) =>
      ruleTypes.has(ruleType)
    ),
    hasUnexpectedBypass: bypassActors.some(
      (actor) => !isExpectedReleaseBotBypass(actor, expectedReleaseBotAppId)
    ),
    name,
  };
}

function rulesetIncludesReleaseProduction(ruleset: unknown): boolean {
  if (!isRecord(ruleset) || !isRecord(ruleset.conditions)) {
    return false;
  }

  const refNameCondition = ruleset.conditions.ref_name;
  if (!isRecord(refNameCondition)) {
    return false;
  }

  const includePatterns = readOptionalArrayProperty(
    refNameCondition,
    "include"
  );
  const includesReleaseProduction = includePatterns.some((pattern) => {
    if (typeof pattern !== "string") {
      return false;
    }

    return (
      pattern === RELEASE_PRODUCTION_REF ||
      pattern === "release/production" ||
      pattern === "~ALL"
    );
  });
  const excludePatterns = readOptionalArrayProperty(
    refNameCondition,
    "exclude"
  );
  const excludesReleaseProduction = excludePatterns.some((pattern) => {
    if (typeof pattern !== "string") {
      return false;
    }

    return (
      pattern === RELEASE_PRODUCTION_REF || pattern === "release/production"
    );
  });

  return includesReleaseProduction && !excludesReleaseProduction;
}

function isExpectedReleaseBotBypass(
  actor: unknown,
  expectedReleaseBotAppId: string
): boolean {
  if (!isRecord(actor)) {
    return false;
  }

  const actorType = readOptionalStringProperty(actor, "actor_type");
  const actorId = readOptionalNumberProperty(actor, "actor_id");

  return (
    actorType === "Integration" &&
    actorId !== undefined &&
    String(actorId) === expectedReleaseBotAppId
  );
}

function isHumanBypass(actor: unknown): boolean {
  if (!isRecord(actor)) {
    return false;
  }

  const actorType = readOptionalStringProperty(actor, "actor_type");
  return (
    actorType === "User" ||
    actorType === "Team" ||
    actorType === "OrganizationAdmin" ||
    actorType === "RepositoryRole"
  );
}

async function githubApiRequest(
  token: string,
  repository: RepositoryCoordinates,
  path: string,
  label: string,
  errors: string[]
): Promise<unknown> {
  const response = await fetch(
    `${GITHUB_API_BASE_URL}/repos/${encodeURIComponent(
      repository.owner
    )}/${encodeURIComponent(repository.repo)}${path}`,
    {
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "User-Agent": "markout-release-governance-audit",
        "X-GitHub-Api-Version": "2022-11-28",
      },
    }
  );
  const responseText = await response.text();

  if (!response.ok) {
    errors.push(
      `Could not read ${label}: GitHub API returned ${response.status}.`
    );
    return undefined;
  }

  if (responseText.trim().length === 0) {
    return {};
  }

  return JSON.parse(responseText) as unknown;
}

function hasRequiredReviewerProtection(environment: unknown): boolean {
  const protectionRules = readOptionalArrayProperty(
    environment,
    "protection_rules"
  );

  return protectionRules.some((protectionRule) => {
    if (
      readOptionalStringProperty(protectionRule, "type") !==
      "required_reviewers"
    ) {
      return false;
    }

    return readOptionalArrayProperty(protectionRule, "reviewers").length > 0;
  });
}

function readToken(environment: Environment, errors: string[]): string {
  const token =
    [
      environment.MARKOUT_GITHUB_TOKEN,
      environment.GH_TOKEN,
      environment.GITHUB_TOKEN,
    ]
      .find((candidate) => isNonEmptyString(candidate))
      ?.trim() ?? "";

  if (token.length === 0) {
    errors.push(
      "Set MARKOUT_GITHUB_TOKEN, GH_TOKEN, or GITHUB_TOKEN with repository settings read access."
    );
  }

  return token;
}

function readRepository(
  environment: Environment,
  errors: string[]
): RepositoryCoordinates | undefined {
  const configuredRepository = environment.GITHUB_REPOSITORY?.trim();
  const repository = isNonEmptyString(configuredRepository)
    ? configuredRepository
    : DEFAULT_REPOSITORY;
  const [owner, repo] = repository.split("/");

  if (!isNonEmptyString(owner) || !isNonEmptyString(repo)) {
    errors.push("GITHUB_REPOSITORY must be in owner/repo format.");
    return undefined;
  }

  return {
    owner,
    repo,
  };
}

function readArrayProperty(
  value: unknown,
  propertyName: string,
  label: string,
  errors: string[]
): unknown[] {
  if (!isRecord(value) || !Array.isArray(value[propertyName])) {
    errors.push(`${label} response must include a ${propertyName} array.`);
    return [];
  }

  return value[propertyName];
}

function readOptionalArrayProperty(
  value: unknown,
  propertyName: string
): unknown[] {
  if (!isRecord(value) || !Array.isArray(value[propertyName])) {
    return [];
  }

  return value[propertyName];
}

function readOptionalNumberProperty(
  value: unknown,
  propertyName: string
): number | undefined {
  if (!isRecord(value) || typeof value[propertyName] !== "number") {
    return undefined;
  }

  return value[propertyName];
}

function readOptionalStringProperty(
  value: unknown,
  propertyName: string
): string | undefined {
  if (!isRecord(value) || typeof value[propertyName] !== "string") {
    return undefined;
  }

  return value[propertyName];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function reportAuditResult(errors: string[], warnings: string[]): void {
  for (const warning of warnings) {
    console.warn(`Warning: ${warning}`);
  }

  if (errors.length > 0) {
    console.error("MarkOut GitHub release governance audit failed:");
    for (const error of errors) {
      console.error(`- ${error}`);
    }
    process.exitCode = 1;
    return;
  }

  console.log("MarkOut GitHub release governance audit passed.");
}

const isDirectExecution =
  process.argv[1]?.endsWith("check-github-release-governance.ts") ?? false;

if (isDirectExecution) {
  void main().catch((error: unknown) => {
    console.error("MarkOut GitHub release governance audit failed.", error);
    process.exitCode = 1;
  });
}
