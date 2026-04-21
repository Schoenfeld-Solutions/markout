import { FakeMailboxItem, installOfficeEnvironment } from "./helpers";

describe("commands", () => {
  beforeEach(() => {
    jest.resetModules();
  });

  it("keeps the commands runtime inert when only taskpane commands remain", async () => {
    const environment = installOfficeEnvironment({
      mailboxItem: new FakeMailboxItem("<div>Original</div>"),
    });
    const commandsModule = await import("../src/commands/commands");

    await environment.triggerReady();

    expect(() => {
      commandsModule.registerCommandHandlers();
    }).not.toThrow();
    expect(Office.actions.associate).not.toHaveBeenCalled();
  });
});
