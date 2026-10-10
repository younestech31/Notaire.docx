export interface DesktopOfficeBridge {
  getDataDirectory(): Promise<string>;
  invoke<T>(command: string, payload: unknown): Promise<T>;
}

export function getDesktopBridge(): DesktopOfficeBridge | null {
  if (typeof window === 'undefined') return null;
  const bridge = (
    window as unknown as { __NOTAIRE_DESKTOP__?: DesktopOfficeBridge }
  ).__NOTAIRE_DESKTOP__;
  return bridge ?? null;
}
