export const MARKOUT_FRAGMENT_HOST_CLASS = "markout-fragment-host";
export const MARKOUT_FRAGMENT_RENDERED_CLASS = "markout-fragment-rendered";
export const MARKOUT_RENDERED_CLASS = "markout-rendered";

export function containsMarkOutFragmentMarker(html: string): boolean {
  const documentFragment = new DOMParser().parseFromString(html, "text/html");

  return (
    documentFragment.body.querySelector(
      `.${MARKOUT_FRAGMENT_HOST_CLASS}, .mo.${MARKOUT_FRAGMENT_RENDERED_CLASS}`
    ) !== null
  );
}

export function containsMarkOutFullRenderMarker(html: string): boolean {
  const documentFragment = new DOMParser().parseFromString(html, "text/html");

  return (
    documentFragment.body.querySelector(`.mo.${MARKOUT_RENDERED_CLASS}`) !==
    null
  );
}
