import { createPreviewManifest } from "../scripts/generate-preview-manifest";

describe("preview manifest generation", () => {
  it("rewrites the beta manifest urls to a preview host and updates the display name", () => {
    const sourceManifest = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp>
  <DisplayName DefaultValue="MarkOut (Beta)" />
  <AppDomains>
    <AppDomain>https://schoenfeld-solutions.github.io</AppDomain>
  </AppDomains>
  <FormSettings>
    <Form>
      <DesktopSettings>
        <SourceLocation DefaultValue="https://schoenfeld-solutions.github.io/markout/outlook-beta/taskpane.html" />
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Resources>
    <bt:Urls>
      <bt:Url id="Commands.Url" DefaultValue="https://schoenfeld-solutions.github.io/markout/outlook-beta/commands.html" />
      <bt:Url id="Taskpane.Url" DefaultValue="https://schoenfeld-solutions.github.io/markout/outlook-beta/taskpane.html" />
      <bt:Url id="WebViewRuntime.Url" DefaultValue="https://schoenfeld-solutions.github.io/markout/outlook-beta/commands.html" />
      <bt:Url id="JSRuntime.Url" DefaultValue="https://schoenfeld-solutions.github.io/markout/outlook-beta/launchevent.js" />
    </bt:Urls>
  </Resources>
</OfficeApp>
`.trim();

    const previewManifest = createPreviewManifest(sourceManifest, {
      displayName: 'MarkOut (Preview PR #12 "QA")',
      previewBaseUrl: "https://pr-12.markout-preview.pages.dev",
    });

    expect(previewManifest).toContain(
      '<DisplayName DefaultValue="MarkOut (Preview PR #12 &quot;QA&quot;)" />'
    );
    expect(previewManifest).toContain(
      "<AppDomain>https://pr-12.markout-preview.pages.dev</AppDomain>"
    );
    expect(previewManifest).toContain(
      'DefaultValue="https://pr-12.markout-preview.pages.dev/outlook-beta/taskpane.html"'
    );
    expect(previewManifest).toContain(
      'DefaultValue="https://pr-12.markout-preview.pages.dev/outlook-beta/launchevent.js"'
    );
  });

  it("fails when the source manifest does not contain the expected pages base url", () => {
    expect(() =>
      createPreviewManifest("<OfficeApp />", {
        displayName: "MarkOut (Preview)",
        previewBaseUrl: "https://preview.example.com",
      })
    ).toThrow(
      "Preview manifest generation expected the source manifest to contain https://schoenfeld-solutions.github.io/markout."
    );
  });
});
