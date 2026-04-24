import {
  checkDeployableManifestInvariants,
  checkOfficeMailManifestInvariants,
  type ManifestSnapshot,
  type RuntimeChannelConfigSnapshot,
} from "../scripts/check-repo-contracts";

const runtimeChannelConfig: RuntimeChannelConfigSnapshot = {
  addInId: "05c2e1c9-3e1d-406e-9a91-e9ac64854143",
  appBaseUrl: "https://schoenfeld-solutions.github.io/markout/outlook",
  channelId: "production",
  commandsUrl:
    "https://schoenfeld-solutions.github.io/markout/outlook/commands.html?channel=production",
  launcheventUrl:
    "https://schoenfeld-solutions.github.io/markout/outlook/launchevent.js?channel=production",
  storageNamespace: "markout.production",
  supportUrl: "https://github.com/Schoenfeld-Solutions/markout",
  taskpaneUrl:
    "https://schoenfeld-solutions.github.io/markout/outlook/taskpane.html?channel=production",
};

function createManifestSnapshot(
  overrides: Partial<ManifestSnapshot> = {}
): ManifestSnapshot {
  const text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xsi:type="MailApp">
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.12" />
    </Sets>
  </Requirements>
  <Permissions>ReadWriteItem</Permissions>
  <ExtensionPoint xsi:type="MessageComposeCommandSurface">
  </ExtensionPoint>
  <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  </ExtensionPoint>
  <ExtensionPoint xsi:type="LaunchEvent">
    <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock" />
    <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="SoftBlock" />
  </ExtensionPoint>
</OfficeApp>`;

  return {
    appDomains: ["https://schoenfeld-solutions.github.io"],
    displayName: "MarkOut",
    highResolutionIconUrl:
      "https://raw.githubusercontent.com/Schoenfeld-Solutions/markout/main/assets/icon-128.png",
    iconUrl:
      "https://raw.githubusercontent.com/Schoenfeld-Solutions/markout/main/assets/icon-64.png",
    id: runtimeChannelConfig.addInId,
    path: "manifest.xml",
    sourceLocation: runtimeChannelConfig.taskpaneUrl,
    supportUrl: runtimeChannelConfig.supportUrl,
    text,
    urls: {
      "Commands.Url": runtimeChannelConfig.commandsUrl,
      "JSRuntime.Url": runtimeChannelConfig.launcheventUrl,
      "Taskpane.Url": runtimeChannelConfig.taskpaneUrl,
      "WebViewRuntime.Url": runtimeChannelConfig.commandsUrl,
    },
    version: "1.0.1.0",
    ...overrides,
  };
}

describe("repository contract manifest checks", () => {
  it("accepts the expected Office mail and deployable manifest invariants", () => {
    const snapshot = createManifestSnapshot();
    const contractErrors: string[] = [];

    checkOfficeMailManifestInvariants(snapshot, contractErrors);
    checkDeployableManifestInvariants(
      runtimeChannelConfig,
      snapshot,
      contractErrors
    );

    expect(contractErrors).toEqual([]);
  });

  it("rejects deployable manifests that leak local channel URLs", () => {
    const snapshot = createManifestSnapshot({
      sourceLocation: "https://localhost:3000/taskpane.html?channel=local",
      text: `${createManifestSnapshot().text}\nhttps://localhost:3000?channel=local`,
    });
    const contractErrors: string[] = [];

    checkDeployableManifestInvariants(
      runtimeChannelConfig,
      snapshot,
      contractErrors
    );

    expect(contractErrors).toEqual(
      expect.arrayContaining([
        "manifest.xml is deployable and must not reference localhost.",
        "manifest.xml is deployable and must not reference the local channel.",
      ])
    );
  });

  it("rejects missing Smart Alerts launch event wiring", () => {
    const snapshot = createManifestSnapshot({
      text: createManifestSnapshot().text.replace(
        '<LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock" />',
        ""
      ),
    });
    const contractErrors: string[] = [];

    checkOfficeMailManifestInvariants(snapshot, contractErrors);

    expect(contractErrors).toContain(
      "manifest.xml is missing message send Smart Alerts handler."
    );
  });

  it("rejects deployable manifest URLs outside the runtime path", () => {
    const snapshot = createManifestSnapshot({
      urls: {
        ...createManifestSnapshot().urls,
        "Taskpane.Url":
          "https://schoenfeld-solutions.github.io/markout/other/taskpane.html?channel=production",
      },
    });
    const contractErrors: string[] = [];

    checkDeployableManifestInvariants(
      runtimeChannelConfig,
      snapshot,
      contractErrors
    );

    expect(contractErrors).toContain(
      "manifest.xml Taskpane.Url must stay under `/markout/outlook/`."
    );
  });
});
