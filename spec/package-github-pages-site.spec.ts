import { mkdir, mkdtemp, readFile, writeFile } from "fs/promises";
import { tmpdir } from "os";
import path from "path";
import { packageGithubPagesSite } from "../scripts/package-github-pages-site";

describe("github pages packaging", () => {
  it("packages production and beta channels from separate source roots", async () => {
    const workspaceRoot = await mkdtemp(path.join(tmpdir(), "markout-pages-"));
    const betaRoot = path.join(workspaceRoot, "beta");
    const productionRoot = path.join(workspaceRoot, "production");
    const outputRoot = path.join(workspaceRoot, "out");

    await writeFixture(betaRoot, "site/index.html", "<html>beta index</html>");
    await writeFixture(betaRoot, "site/404.html", "<html>beta 404</html>");
    await writeFixture(betaRoot, "manifest.beta.xml", "<beta />");
    await writeFixture(betaRoot, "manifest.xml", "<wrong-beta-production />");
    await writeFixture(betaRoot, "assets/icon.txt", "beta asset");
    await writeFixture(betaRoot, "dist/taskpane.html", "beta taskpane");
    await writeFixture(betaRoot, "dist/commands.js", "beta commands");

    await writeFixture(
      productionRoot,
      "manifest.xml",
      "<production-manifest />"
    );
    await writeFixture(
      productionRoot,
      "manifest.beta.xml",
      "<wrong-production-beta />"
    );
    await writeFixture(productionRoot, "assets/icon.txt", "production asset");
    await writeFixture(
      productionRoot,
      "dist/taskpane.html",
      "production taskpane"
    );
    await writeFixture(
      productionRoot,
      "dist/commands.js",
      "production commands"
    );

    await packageGithubPagesSite({
      betaRoot,
      outputRoot,
      productionRoot,
    });

    await expectFile(outputRoot, "index.html", "<html>beta index</html>");
    await expectFile(outputRoot, "404.html", "<html>beta 404</html>");
    await expectFile(outputRoot, "manifest.xml", "<production-manifest />");
    await expectFile(outputRoot, "manifest.beta.xml", "<beta />");
    await expectFile(outputRoot, "assets/icon.txt", "beta asset");
    await expectFile(
      outputRoot,
      "outlook/taskpane.html",
      "production taskpane"
    );
    await expectFile(outputRoot, "outlook/commands.js", "production commands");
    await expectFile(outputRoot, "outlook/assets/icon.txt", "production asset");
    await expectFile(outputRoot, "outlook-beta/taskpane.html", "beta taskpane");
    await expectFile(outputRoot, "outlook-beta/commands.js", "beta commands");
    await expectFile(outputRoot, "outlook-beta/assets/icon.txt", "beta asset");
  });
});

async function expectFile(
  rootDirectory: string,
  relativePath: string,
  expectedContent: string
): Promise<void> {
  await expect(
    readFile(path.join(rootDirectory, relativePath), "utf8")
  ).resolves.toBe(expectedContent);
}

async function writeFixture(
  rootDirectory: string,
  relativePath: string,
  content: string
): Promise<void> {
  const targetPath = path.join(rootDirectory, relativePath);
  await mkdir(path.dirname(targetPath), { recursive: true });
  await writeFile(targetPath, content, { encoding: "utf8", flag: "w" });
}
