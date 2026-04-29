import {
  auditRuleset,
  validateReleaseProductionRulesets,
} from "../scripts/check-github-release-governance";

const RELEASE_BOT_APP_ID = "12345";

function createReleaseProductionRuleset(
  overrides: Partial<{
    bypassActors: unknown[];
    enforcement: string;
    exclude: string[];
    include: string[];
    rules: unknown[];
    target: string;
  }> = {}
): unknown {
  return {
    bypass_actors: overrides.bypassActors ?? [
      {
        actor_id: Number(RELEASE_BOT_APP_ID),
        actor_type: "Integration",
      },
    ],
    conditions: {
      ref_name: {
        exclude: overrides.exclude ?? [],
        include: overrides.include ?? ["refs/heads/release/production"],
      },
    },
    enforcement: overrides.enforcement ?? "active",
    name: "protect release production",
    rules: overrides.rules ?? [
      { type: "creation" },
      { type: "deletion" },
      { type: "update" },
    ],
    target: overrides.target ?? "branch",
  };
}

describe("github release governance audit", () => {
  it("accepts an active release production ruleset with only the release bot bypass", () => {
    const errors = validateReleaseProductionRulesets(
      [createReleaseProductionRuleset()],
      RELEASE_BOT_APP_ID
    );

    expect(errors).toEqual([]);
    expect(
      auditRuleset(createReleaseProductionRuleset(), RELEASE_BOT_APP_ID)
    ).toMatchObject({
      hasAutomationBypass: true,
      hasHumanBypass: false,
      hasReleaseProductionCondition: true,
      hasRequiredRuleTypes: true,
      hasUnexpectedBypass: false,
    });
  });

  it("rejects missing rule types and inactive rulesets", () => {
    expect(
      validateReleaseProductionRulesets(
        [
          createReleaseProductionRuleset({
            rules: [{ type: "creation" }, { type: "update" }],
          }),
        ],
        RELEASE_BOT_APP_ID
      )
    ).toContain(
      "release/production must have an active ruleset that blocks creation, update, and deletion."
    );

    expect(
      validateReleaseProductionRulesets(
        [
          createReleaseProductionRuleset({
            enforcement: "disabled",
          }),
        ],
        RELEASE_BOT_APP_ID
      )
    ).toContain(
      "release/production must have an active ruleset that blocks creation, update, and deletion."
    );
  });

  it("rejects release production rulesets excluded from all-branch patterns", () => {
    expect(
      auditRuleset(
        createReleaseProductionRuleset({
          exclude: ["refs/heads/release/production"],
          include: ["~ALL"],
        }),
        RELEASE_BOT_APP_ID
      )
    ).toMatchObject({
      hasReleaseProductionCondition: false,
    });

    expect(
      validateReleaseProductionRulesets(
        [
          createReleaseProductionRuleset({
            exclude: ["refs/heads/release/production"],
            include: ["~ALL"],
          }),
        ],
        RELEASE_BOT_APP_ID
      )
    ).toContain(
      "release/production must have an active ruleset that blocks creation, update, and deletion."
    );
  });

  it("rejects missing, foreign, human, and mixed bypass actors", () => {
    expect(
      validateReleaseProductionRulesets(
        [
          createReleaseProductionRuleset({
            bypassActors: [],
          }),
        ],
        RELEASE_BOT_APP_ID
      )
    ).toContain(
      "release/production ruleset must allow the markout-release-bot GitHub App as its only automation bypass."
    );

    expect(
      validateReleaseProductionRulesets(
        [
          createReleaseProductionRuleset({
            bypassActors: [{ actor_id: 99999, actor_type: "Integration" }],
          }),
        ],
        RELEASE_BOT_APP_ID
      )
    ).toEqual([
      "release/production ruleset must allow the markout-release-bot GitHub App as its only automation bypass.",
      "release/production ruleset must not include bypass actors other than the markout-release-bot GitHub App.",
    ]);

    expect(
      validateReleaseProductionRulesets(
        [
          createReleaseProductionRuleset({
            bypassActors: [{ actor_id: 42, actor_type: "User" }],
          }),
        ],
        RELEASE_BOT_APP_ID
      )
    ).toEqual([
      "release/production ruleset must allow the markout-release-bot GitHub App as its only automation bypass.",
      "release/production ruleset must not include User, Team, or admin bypass actors.",
      "release/production ruleset must not include bypass actors other than the markout-release-bot GitHub App.",
    ]);

    expect(
      validateReleaseProductionRulesets(
        [
          createReleaseProductionRuleset({
            bypassActors: [
              {
                actor_id: Number(RELEASE_BOT_APP_ID),
                actor_type: "Integration",
              },
              { actor_id: 99999, actor_type: "Integration" },
            ],
          }),
        ],
        RELEASE_BOT_APP_ID
      )
    ).toContain(
      "release/production ruleset must not include bypass actors other than the markout-release-bot GitHub App."
    );
  });
});
