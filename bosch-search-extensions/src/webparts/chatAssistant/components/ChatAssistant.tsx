import * as React from 'react';
import { useState, useRef, useEffect, useCallback } from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useGraphClient } from '../../../hooks/useGraphClient';
import { CopilotChatService } from '../../../services/CopilotChatService';
import { CopilotDetectionService } from '../../../services/CopilotDetectionService';
import { GraphSearchService } from '../../../services/GraphSearchService';
import { IChatMessage, IChatCitation } from '../../../models';
import { M365_COPILOT_SEARCH_URL } from '../../../common/Constants';
import styles from './ChatAssistant.module.scss';

export interface IChatAssistantProps {
  context: WebPartContext;
  groundingMode: 'work' | 'web' | 'both';
  maxConversationTurns: number;
  panelMode: 'sidePanel' | 'inline';
  welcomeMessage: string;
  suggestedQuestions: string[];
  showCopilotLink: boolean;
}

export const ChatAssistant: React.FC<IChatAssistantProps> = ({
  context,
  groundingMode,
  maxConversationTurns,
  panelMode,
  welcomeMessage,
  suggestedQuestions,
  showCopilotLink,
}) => {
  const { graphClient } = useGraphClient(context);
  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [messages, setMessages] = useState<IChatMessage[]>([]);
  const [inputValue, setInputValue] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [hasCopilot, setHasCopilot] = useState<boolean | null>(null);
  const [conversationId, setConversationId] = useState<string | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const abortRef = useRef(false);

  const scrollToBottom = useCallback(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, []);

  useEffect(() => {
    scrollToBottom();
  }, [messages, scrollToBottom]);

  // Check Copilot license on mount
  useEffect(() => {
    if (!graphClient) return;
    const detector = new CopilotDetectionService(graphClient);
    detector.hasCopilotLicense().then(setHasCopilot);
  }, [graphClient]);

  const handleSend = async (messageText?: string): Promise<void> => {
    const text = messageText || inputValue.trim();
    if (!text || !graphClient || hasCopilot === null || isGenerating) return;

    setInputValue('');
    abortRef.current = false;

    const userMessage: IChatMessage = {
      role: 'user',
      content: text,
      timestamp: new Date(),
    };

    const updatedMessages = [...messages, userMessage];
    const trimmedMessages = updatedMessages.slice(-maxConversationTurns * 2);
    setMessages(trimmedMessages);
    setIsGenerating(true);

    try {
      if (hasCopilot) {
        await handleCopilotChat(text);
      } else {
        await handleGraphSearchChat(text);
      }
    } catch (err) {
      if (!abortRef.current) {
        console.error('[ChatAssistant] Error:', err);
        const errorMessage: IChatMessage = {
          role: 'assistant',
          content: `Sorry, I encountered an error: ${err instanceof Error ? err.message : 'Unknown error'}`,
          timestamp: new Date(),
        };
        setMessages((prev) => [...prev, errorMessage]);
      }
    } finally {
      if (!abortRef.current) {
        setIsGenerating(false);
      }
    }
  };

  /**
   * Copilot-licensed path: multi-turn Copilot Chat API conversation.
   * Creates a conversation on first message, reuses it for subsequent turns.
   */
  const handleCopilotChat = async (text: string): Promise<void> => {
    const chatService = new CopilotChatService(graphClient!);
    const enableWebGrounding = groundingMode === 'web' || groundingMode === 'both';

    let currentConversationId = conversationId;

    // Create conversation on first message
    if (!currentConversationId) {
      const conversation = await chatService.createConversation();
      if (abortRef.current) return;
      currentConversationId = conversation.id;
      setConversationId(currentConversationId);
    }

    const result = await chatService.sendMessage(currentConversationId, text, {
      enableWebGrounding,
    });
    if (abortRef.current) return;

    const citations: IChatCitation[] = result.attributions
      .filter((a) => a.url)
      .map((attr) => ({
        title: attr.title || 'Source',
        url: attr.url || '',
      }));

    const assistantMessage: IChatMessage = {
      role: 'assistant',
      content: result.responseText,
      timestamp: new Date(),
      citations,
    };

    setMessages((prev) => [...prev, assistantMessage]);
  };

  /**
   * Non-Copilot path: uses Graph Search API to find relevant results
   * and displays them as a structured response (no LLM synthesis).
   */
  const handleGraphSearchChat = async (text: string): Promise<void> => {
    const searchService = new GraphSearchService(graphClient!);
    const searchResponse = await searchService.search(
      text,
      ['externalItem', 'driveItem', 'listItem'],
      0,
      5
    );
    if (abortRef.current) return;

    const citations: IChatCitation[] = searchResponse.results.map((r) => ({
      title: r.title,
      url: r.url,
    }));

    let responseText: string;
    if (searchResponse.results.length === 0) {
      responseText = 'No results found for your query. Try rephrasing or using different keywords.';
    } else {
      responseText = searchResponse.results
        .map((r, i) => `**[${i + 1}] ${r.title}**\n${r.summary || 'No summary available'}`)
        .join('\n\n');
    }

    const assistantMessage: IChatMessage = {
      role: 'assistant',
      content: responseText,
      timestamp: new Date(),
      citations,
    };

    setMessages((prev) => [...prev, assistantMessage]);
  };

  const handleClear = (): void => {
    abortRef.current = true;
    setMessages([]);
    setConversationId(null);
    setIsGenerating(false);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const chatContent = (
    <div className={panelMode === 'inline' ? styles.inlineContainer : styles.chatContainer}>
      <div className={styles.chatHeader}>
        <span className={styles.chatTitle}>
          {hasCopilot ? 'Copilot Chat' : 'Search Assistant'}
        </span>
        {hasCopilot && <span className={styles.copilotBadge}>Copilot</span>}
        {!hasCopilot && hasCopilot !== null && (
          <span className={styles.copilotBadge} style={{ background: '#107c10' }}>
            Graph Search
          </span>
        )}
        <button className={styles.clearButton} onClick={handleClear}>
          Clear Chat
        </button>
      </div>

      <div className={styles.messagesArea}>
        {messages.length === 0 && (
          <>
            {welcomeMessage && (
              <div className={styles.welcomeMessage}>{welcomeMessage}</div>
            )}
            {suggestedQuestions.length > 0 && (
              <div className={styles.suggestedQuestions}>
                {suggestedQuestions.map((q, i) => (
                  <button
                    key={i}
                    className={styles.suggestedButton}
                    onClick={() => handleSend(q)}
                  >
                    {q}
                  </button>
                ))}
              </div>
            )}
          </>
        )}

        {messages.map((msg, i) => (
          <div
            key={i}
            className={`${styles.messageBubble} ${
              msg.role === 'user' ? styles.userMessage : styles.assistantMessage
            }`}
          >
            {msg.content}
            {msg.isStreaming && <span className={styles.streamingCursor} />}
            {msg.citations && msg.citations.length > 0 && !msg.isStreaming && (
              <div className={styles.messageCitations}>
                {msg.citations.map((cite, ci) => (
                  <a
                    key={ci}
                    href={cite.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className={styles.citationLink}
                  >
                    [{ci + 1}] {cite.title}
                  </a>
                ))}
              </div>
            )}
          </div>
        ))}

        {isGenerating && (
          <div className={styles.assistantMessage}>
            <Spinner label={hasCopilot ? 'Asking Copilot...' : 'Searching...'} />
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      <div className={styles.inputArea}>
        <input
          type="text"
          className={styles.inputField}
          placeholder={hasCopilot ? 'Ask Copilot...' : 'Ask a question...'}
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={isGenerating || hasCopilot === null}
        />
        <button
          className={styles.sendButton}
          onClick={() => handleSend()}
          disabled={isGenerating || !inputValue.trim()}
          title="Send"
        >
          <Icon iconName="Send" />
        </button>
      </div>

      {showCopilotLink && hasCopilot && messages.length > 0 && !isGenerating && (
        <div style={{ padding: '8px 16px', textAlign: 'right', borderTop: '1px solid #edebe9' }}>
          <Link href={M365_COPILOT_SEARCH_URL} target="_blank">
            <Icon iconName="OpenInNewWindow" style={{ marginRight: 4 }} />
            Open in M365 Copilot
          </Link>
        </div>
      )}
    </div>
  );

  if (hasCopilot === null) {
    return (
      <div className={styles.configWarning}>
        <Spinner label="Checking Copilot license..." />
      </div>
    );
  }

  if (panelMode === 'inline') {
    return chatContent;
  }

  // Side panel mode
  return (
    <>
      <button
        className={styles.chatTrigger}
        onClick={() => setIsPanelOpen(true)}
        title="Open Search Assistant"
        aria-label="Open Search Assistant"
      >
        <Icon iconName="Chat" className={styles.triggerIcon} />
      </button>

      <Panel
        isOpen={isPanelOpen}
        onDismiss={() => setIsPanelOpen(false)}
        type={PanelType.medium}
        headerText=""
        isLightDismiss
        styles={{
          content: { padding: 0, height: '100%' },
          scrollableContent: { height: '100%' },
        }}
      >
        {chatContent}
      </Panel>
    </>
  );
};
