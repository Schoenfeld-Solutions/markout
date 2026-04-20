export class Debounce {
  private triggerHandle: ReturnType<typeof setTimeout> | null = null;

  public constructor(
    private readonly action: () => void | Promise<void>,
    private readonly delay = 500
  ) {}

  public trigger(): void {
    if (this.triggerHandle !== null) {
      clearTimeout(this.triggerHandle);
    }

    this.triggerHandle = setTimeout(() => {
      void this.action();
      this.triggerHandle = null;
    }, this.delay);
  }
}
