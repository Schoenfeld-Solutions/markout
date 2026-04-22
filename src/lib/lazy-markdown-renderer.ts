import type { MarkdownRenderer } from "./renderer";

export function createLazyMarkdownRenderer(): MarkdownRenderer {
  let rendererPromise: Promise<MarkdownRenderer> | null = null;

  return {
    render: async (options) => {
      rendererPromise ??= import(
        /* webpackChunkName: "markout-renderer" */ "./renderer"
      )
        .then(({ createMarkdownRenderer }) => createMarkdownRenderer())
        .catch((error: unknown) => {
          rendererPromise = null;
          throw error;
        });

      return (await rendererPromise).render(options);
    },
  };
}
