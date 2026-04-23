import type { ReactElement, ReactNode } from "react";

function ToolbarIcon({ children }: { children: ReactNode }): ReactElement {
  return (
    <svg
      aria-hidden="true"
      fill="none"
      height="18"
      stroke="currentColor"
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="1.6"
      viewBox="0 0 20 20"
      width="18"
    >
      {children}
    </svg>
  );
}

export function InsertIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M10 3v9" />
      <path d="M7 9.5 10 12.5l3-3" />
      <path d="M4 15.5h12" />
    </ToolbarIcon>
  );
}

export function InfoIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <circle cx="10" cy="10" r="6.5" />
      <path d="M10 8v4" />
      <path d="M10 6.2h.01" />
    </ToolbarIcon>
  );
}

export function HelpIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M7.8 7.8a2.2 2.2 0 1 1 3.5 1.7c-.8.6-1.3 1-1.3 2" />
      <path d="M10 14.3h.01" />
      <circle cx="10" cy="10" r="6.5" />
    </ToolbarIcon>
  );
}

export function CreditsIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="m10 4.3 1.5 3.1 3.4.5-2.5 2.5.6 3.5-3-1.6-3 1.6.6-3.5L5 7.9l3.5-.5Z" />
    </ToolbarIcon>
  );
}

export function SettingsIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <circle cx="10" cy="10" r="2.2" />
      <path d="M10 4.5v1.4" />
      <path d="M10 14.1v1.4" />
      <path d="m5.8 5.8 1 1" />
      <path d="m13.2 13.2 1 1" />
      <path d="M4.5 10h1.4" />
      <path d="M14.1 10h1.4" />
      <path d="m5.8 14.2 1-1" />
      <path d="m13.2 6.8 1-1" />
    </ToolbarIcon>
  );
}

export function DeveloperIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="m7.5 6-3 4 3 4" />
      <path d="m12.5 6 3 4-3 4" />
      <path d="m11 5-2 10" />
    </ToolbarIcon>
  );
}

export function RepositoryIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M4.5 6.5A1.5 1.5 0 0 1 6 5h8a1.5 1.5 0 0 1 1.5 1.5v7A1.5 1.5 0 0 1 14 15H6a1.5 1.5 0 0 1-1.5-1.5Z" />
      <path d="M7 8.2h6" />
      <path d="M7 10.5h4" />
      <path d="M7 12.8h5" />
    </ToolbarIcon>
  );
}

export function DocsIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M6.5 4.8h5.8l2.2 2.2v8.2H6.5Z" />
      <path d="M12.3 4.8v2.4h2.2" />
      <path d="M8.4 10h4.8" />
      <path d="M8.4 12.4h3.6" />
    </ToolbarIcon>
  );
}

export function CompanyIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="M6 15.2V6.2h8v9" />
      <path d="M4.5 15.2h11" />
      <path d="M8.2 8.4h1.2" />
      <path d="M10.6 8.4h1.2" />
      <path d="M8.2 10.8h1.2" />
      <path d="M10.6 10.8h1.2" />
      <path d="M9.3 15.2v-2.5h1.4v2.5" />
    </ToolbarIcon>
  );
}

export function UpstreamIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <path d="m6.5 9.5 3-3 3 3" />
      <path d="M9.5 6.8v6.7" />
      <path d="M5.5 13.8h8" />
    </ToolbarIcon>
  );
}

export function ForkIcon(): ReactElement {
  return (
    <ToolbarIcon>
      <circle cx="6.5" cy="5.8" r="1.2" />
      <circle cx="13.5" cy="5.8" r="1.2" />
      <circle cx="10" cy="14.1" r="1.2" />
      <path d="M6.5 7v2.1c0 1.2 1 2.2 2.2 2.2H10" />
      <path d="M13.5 7v2.1c0 1.2-1 2.2-2.2 2.2H10" />
      <path d="M10 11.3v1.6" />
    </ToolbarIcon>
  );
}

export function IntroComposeIllustration(): ReactElement {
  return (
    <svg
      aria-hidden="true"
      fill="none"
      height="100%"
      viewBox="0 0 240 140"
      width="100%"
    >
      <rect
        fill="currentColor"
        fillOpacity="0.08"
        height="104"
        rx="18"
        width="208"
        x="16"
        y="18"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.14"
        height="72"
        rx="14"
        width="84"
        x="30"
        y="34"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.1"
        height="10"
        rx="5"
        width="84"
        x="128"
        y="38"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.1"
        height="10"
        rx="5"
        width="66"
        x="128"
        y="58"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.1"
        height="10"
        rx="5"
        width="74"
        x="128"
        y="78"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.16"
        height="16"
        rx="8"
        width="48"
        x="128"
        y="100"
      />
    </svg>
  );
}

export function IntroInsertIllustration(): ReactElement {
  return (
    <svg
      aria-hidden="true"
      fill="none"
      height="100%"
      viewBox="0 0 240 140"
      width="100%"
    >
      <rect
        fill="currentColor"
        fillOpacity="0.08"
        height="104"
        rx="18"
        width="208"
        x="16"
        y="18"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.1"
        height="62"
        rx="12"
        width="120"
        x="32"
        y="38"
      />
      <path
        d="M92 50v18"
        stroke="currentColor"
        strokeLinecap="round"
        strokeWidth="6"
      />
      <path
        d="m82 60 10 10 10-10"
        stroke="currentColor"
        strokeLinecap="round"
        strokeLinejoin="round"
        strokeWidth="6"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.16"
        height="62"
        rx="12"
        width="52"
        x="164"
        y="38"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.08"
        height="10"
        rx="5"
        width="32"
        x="174"
        y="52"
      />
      <rect
        fill="currentColor"
        fillOpacity="0.08"
        height="10"
        rx="5"
        width="24"
        x="174"
        y="72"
      />
    </svg>
  );
}
