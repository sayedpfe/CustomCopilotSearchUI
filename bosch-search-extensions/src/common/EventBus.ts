type EventCallback = (data: unknown) => void;

/**
 * Simple pub/sub event bus for cross-web-part communication.
 * Singleton shared across all web parts on the same page.
 */
export class EventBus {
  private static listeners: Map<string, EventCallback[]> = new Map();

  public static on(event: string, callback: EventCallback): void {
    const existing = EventBus.listeners.get(event) || [];
    existing.push(callback);
    EventBus.listeners.set(event, existing);
  }

  public static off(event: string, callback: EventCallback): void {
    const existing = EventBus.listeners.get(event);
    if (!existing) return;
    EventBus.listeners.set(
      event,
      existing.filter((cb) => cb !== callback)
    );
  }

  public static emit(event: string, data: unknown): void {
    const callbacks = EventBus.listeners.get(event);
    if (!callbacks) return;
    callbacks.forEach((cb) => {
      try {
        cb(data);
      } catch (err) {
        console.error(`[EventBus] Error in listener for "${event}":`, err);
      }
    });
  }

  public static clear(): void {
    EventBus.listeners.clear();
  }
}
