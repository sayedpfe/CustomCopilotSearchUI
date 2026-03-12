import { MSGraphClientV3 } from '@microsoft/sp-http';
import { COPILOT_CONVERSATIONS_ENDPOINT } from '../common/Constants';

export interface ICopilotAttribution {
  title?: string;
  url?: string;
}

export interface ICopilotResponseMessage {
  id: string;
  text: string;
  createdDateTime: string;
  attributions?: ICopilotAttribution[];
}

export interface ICopilotConversation {
  id: string;
  createdDateTime: string;
  displayName: string;
  status: string;
  turnCount: number;
  messages?: ICopilotResponseMessage[];
}

export interface IStreamCallbacks {
  onChunk: (textSoFar: string) => void;
  onDone: (fullText: string, attributions: ICopilotAttribution[]) => void;
  onError: (error: Error) => void;
}

/**
 * Service for the Microsoft 365 Copilot Chat API (Preview).
 *
 * Endpoints:
 *   POST /beta/copilot/conversations                       — Create a conversation
 *   POST /beta/copilot/conversations/{id}/chat              — Send a message (sync)
 *   POST /beta/copilot/conversations/{id}/chatOverStream    — Send a message (streaming SSE)
 *
 * Uses enterprise search grounding + web search grounding by default.
 * Returns fully synthesized answers grounded in M365 data.
 *
 * Requires Copilot license.
 * Permissions: Sites.Read.All, Mail.Read, People.Read.All, Chat.Read,
 *              ExternalItem.Read.All, OnlineMeetingTranscript.Read.All
 */
export class CopilotChatService {
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  /**
   * Creates a new Copilot conversation.
   * Returns the conversation object with an id for subsequent chat calls.
   */
  public async createConversation(): Promise<ICopilotConversation> {
    console.log(`[CopilotChatService] POST ${COPILOT_CONVERSATIONS_ENDPOINT} (beta) — Creating conversation`);
    const response = await this.graphClient
      .api(COPILOT_CONVERSATIONS_ENDPOINT)
      .version('beta')
      .post({});
    console.log('[CopilotChatService] Conversation created:', response.id);

    return {
      id: response.id,
      createdDateTime: response.createdDateTime,
      displayName: response.displayName || '',
      status: response.status || 'active',
      turnCount: response.turnCount || 0,
    };
  }

  /**
   * Sends a message to a Copilot conversation and returns the response.
   * This is a synchronous (non-streaming) call.
   *
   * @param conversationId - The conversation ID from createConversation()
   * @param message - The user's natural language message
   * @param options - Optional configuration for grounding and context
   */
  public async sendMessage(
    conversationId: string,
    message: string,
    options?: {
      enableWebGrounding?: boolean;
      contextualResources?: Array<{
        contentType: string;
        contentUrl: string;
      }>;
    }
  ): Promise<{
    responseText: string;
    attributions: ICopilotAttribution[];
    turnCount: number;
  }> {
    const body = this.buildChatBody(message, options);

    console.log(`[CopilotChatService] POST ${COPILOT_CONVERSATIONS_ENDPOINT}/${conversationId}/chat (beta)`, JSON.stringify(body, null, 2));
    const response = await this.graphClient
      .api(`${COPILOT_CONVERSATIONS_ENDPOINT}/${conversationId}/chat`)
      .version('beta')
      .post(body);
    console.log('[CopilotChatService] Chat response received, messages:', response.messages?.length);

    // Extract the assistant's response message (last ResponseMessage in array)
    const messages = response.messages || [];
    const responseMessages = messages.filter(
      (m: Record<string, unknown>) =>
        (m['@odata.type'] as string)?.includes('ResponseMessage')
    );
    // The last ResponseMessage contains the Copilot answer
    const assistantMessage = responseMessages.length > 0
      ? responseMessages[responseMessages.length - 1]
      : undefined;

    const attributions: ICopilotAttribution[] = (assistantMessage?.attributions || [])
      .filter((attr: Record<string, unknown>) =>
        attr.attributionType === 'citation' && attr.seeMoreWebUrl
      )
      .map((attr: Record<string, unknown>) => ({
        title: (attr.providerDisplayName as string) || '',
        url: (attr.seeMoreWebUrl as string) || '',
      }));

    return {
      responseText: assistantMessage?.body?.content || assistantMessage?.text || 'No response generated.',
      attributions,
      turnCount: response.turnCount || 0,
    };
  }

  /**
   * Builds the request body for chat / chatOverStream.
   */
  private buildChatBody(
    message: string,
    options?: {
      enableWebGrounding?: boolean;
      contextualResources?: Array<{ contentType: string; contentUrl: string }>;
    }
  ): Record<string, unknown> {
    const body: Record<string, unknown> = {
      message: { text: message },
      locationHint: { timeZone: 'Europe/Berlin' },
    };

    const ctx: Record<string, unknown> = {};
    let hasCtx = false;

    if (options?.enableWebGrounding === false) {
      ctx.webContext = { isWebEnabled: false };
      hasCtx = true;
    }
    if (options?.contextualResources?.length) {
      ctx.files = options.contextualResources.map((r) => ({ uri: r.contentUrl }));
      hasCtx = true;
    }
    if (hasCtx) body.contextualResources = ctx;

    return body;
  }

  /**
   * Sends a message using the streaming endpoint (/chatOverStream).
   * Text chunks arrive via SSE and are forwarded through callbacks.
   *
   * The MSGraphClientV3 `.responseType('raw')` returns the native fetch Response,
   * allowing us to read the SSE stream from `response.body`.
   */
  public async sendMessageStream(
    conversationId: string,
    message: string,
    callbacks: IStreamCallbacks,
    options?: {
      enableWebGrounding?: boolean;
      contextualResources?: Array<{ contentType: string; contentUrl: string }>;
    }
  ): Promise<void> {
    // chatOverStream uses the same body shape as /chat
    const body = this.buildChatBody(message, options);
    const endpoint = `${COPILOT_CONVERSATIONS_ENDPOINT}/${conversationId}/chatOverStream`;

    console.log(`[CopilotChatService] POST ${endpoint} (beta, streaming)`, JSON.stringify(body, null, 2));

    try {
      // Get raw Response so we can read the SSE stream
      const rawResponse = await this.graphClient
        .api(endpoint)
        .version('beta')
        .responseType('raw' as never)
        .post(body);

      const response = rawResponse as Response;

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Copilot streaming failed (${response.status}): ${errorText}`);
      }

      if (!response.body) {
        throw new Error('Response body is null — streaming not supported by this browser.');
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let fullText = '';
      let attributions: ICopilotAttribution[] = [];
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });

        // Process complete SSE lines
        const lines = buffer.split('\n');
        buffer = lines.pop() || ''; // Keep incomplete line in buffer

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed || !trimmed.startsWith('data:')) continue;

          const data = trimmed.slice(5).trim();
          if (data === '[DONE]') continue;

          try {
            const parsed = JSON.parse(data) as Record<string, unknown>;

            // chatOverStream returns conversation objects with a messages[] array.
            // Each SSE event carries the full cumulative text in messages[0].text.
            const messages = parsed.messages as Array<Record<string, unknown>> | undefined;
            const firstMsg = messages?.[0];
            if (firstMsg) {
              const text = firstMsg.text as string | undefined;
              if (text) {
                fullText = text;
                callbacks.onChunk(fullText);
              }

              // Extract attributions when present on the message
              const rawAttrs = firstMsg.attributions as Array<Record<string, unknown>> | undefined;
              if (rawAttrs?.length) {
                attributions = rawAttrs
                  .filter((a) => a.url || a.seeMoreWebUrl)
                  .map((a) => ({
                    title: (a.providerDisplayName as string) || (a.title as string) || '',
                    url: (a.seeMoreWebUrl as string) || (a.url as string) || '',
                  }));
              }
            }
          } catch {
            // Skip malformed JSON lines
          }
        }
      }

      callbacks.onDone(fullText, attributions);
    } catch (err) {
      callbacks.onError(err instanceof Error ? err : new Error(String(err)));
    }
  }

  /**
   * Convenience method: creates a conversation and sends a single message.
   * Useful for one-shot Q&A (the AI Answer Panel pattern).
   */
  public async askSingleQuestion(
    question: string,
    enableWebGrounding: boolean = false
  ): Promise<{
    responseText: string;
    attributions: ICopilotAttribution[];
    conversationId: string;
  }> {
    const conversation = await this.createConversation();
    const result = await this.sendMessage(conversation.id, question, {
      enableWebGrounding,
    });
    return {
      ...result,
      conversationId: conversation.id,
    };
  }

  /**
   * Streaming version of askSingleQuestion.
   * Creates a conversation and streams the response via callbacks.
   * Text appears word-by-word in ~2-3s instead of waiting 30-40s.
   */
  public async askSingleQuestionStream(
    question: string,
    callbacks: IStreamCallbacks,
    enableWebGrounding: boolean = false
  ): Promise<string> {
    const conversation = await this.createConversation();
    await this.sendMessageStream(conversation.id, question, callbacks, {
      enableWebGrounding,
    });
    return conversation.id;
  }
}
