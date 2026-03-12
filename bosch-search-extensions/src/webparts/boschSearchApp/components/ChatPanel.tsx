import * as React from 'react';
import { useState, useRef, useEffect, useCallback } from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { CopilotChatService, ICopilotAttribution } from '../../../services/CopilotChatService';
import { GraphSearchService } from '../../../services/GraphSearchService';
import { IChatMessage, IChatCitation } from '../../../models';
import { markdownToHtml } from '../../../common/Utils';
import styles from './BoschSearchApp.module.scss';

export interface IChatPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  onOpen: () => void;
  graphClient: MSGraphClientV3 | undefined;
  groundingMode: 'work' | 'web' | 'both';
  hasCopilot: boolean | null;
}

export const ChatPanel: React.FC<IChatPanelProps> = ({
  isOpen,
  onDismiss,
  onOpen,
  graphClient,
  groundingMode,
  hasCopilot,
}) => {
  const [messages, setMessages] = useState<IChatMessage[]>([]);
  const [inputValue, setInputValue] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [conversationId, setConversationId] = useState<string | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const abortRef = useRef(false);

  const scrollToBottom = useCallback(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, []);

  useEffect(() => {
    scrollToBottom();
  }, [messages, scrollToBottom]);

  const handleSend = async (messageText?: string): Promise<void> => {
    const text = messageText || inputValue.trim();
    if (!text || !graphClient || isGenerating) return;

    setInputValue('');
    abortRef.current = false;

    const userMessage: IChatMessage = {
      role: 'user',
      content: text,
      timestamp: new Date(),
    };
    setMessages((prev) => [...prev, userMessage]);
    setIsGenerating(true);

    try {
      if (hasCopilot) {
        const chatService = new CopilotChatService(graphClient);
        const enableWeb = groundingMode === 'web' || groundingMode === 'both';

        let currentConversationId = conversationId;
        if (!currentConversationId) {
          const conversation = await chatService.createConversation();
          if (abortRef.current) return;
          currentConversationId = conversation.id;
          setConversationId(currentConversationId);
        }

        // Add a placeholder assistant message that we'll update with stream chunks
        const placeholderIndex = messages.length + 1; // +1 for the user message we just added
        setMessages((prev) => [
          ...prev,
          { role: 'assistant', content: '', timestamp: new Date() },
        ]);

        await chatService.sendMessageStream(
          currentConversationId,
          text,
          {
            onChunk: (textSoFar: string) => {
              if (abortRef.current) return;
              // Update the last (assistant) message in-place
              setMessages((prev) => {
                const updated = [...prev];
                const last = updated[updated.length - 1];
                if (last && last.role === 'assistant') {
                  updated[updated.length - 1] = { ...last, content: textSoFar };
                }
                return updated;
              });
            },
            onDone: (fullText: string, attributions: ICopilotAttribution[]) => {
              if (abortRef.current) return;
              const citations: IChatCitation[] = attributions
                .filter((a) => a.url)
                .map((attr) => ({ title: attr.title || 'Source', url: attr.url || '' }));
              setMessages((prev) => {
                const updated = [...prev];
                const last = updated[updated.length - 1];
                if (last && last.role === 'assistant') {
                  updated[updated.length - 1] = { ...last, content: fullText, citations };
                }
                return updated;
              });
            },
            onError: (err: Error) => {
              if (!abortRef.current) {
                setMessages((prev) => {
                  const updated = [...prev];
                  const last = updated[updated.length - 1];
                  if (last && last.role === 'assistant') {
                    updated[updated.length - 1] = { ...last, content: `Error: ${err.message}` };
                  }
                  return updated;
                });
              }
            },
          },
          { enableWebGrounding: enableWeb }
        );
      } else {
        const searchService = new GraphSearchService(graphClient);
        const response = await searchService.search(text, ['externalItem', 'driveItem', 'listItem'], 0, 5);
        if (abortRef.current) return;

        const citations: IChatCitation[] = response.results.map((r) => ({
          title: r.title,
          url: r.url,
        }));

        const responseText =
          response.results.length === 0
            ? 'No results found. Try different keywords.'
            : response.results
                .map((r, i) => `**[${i + 1}] ${r.title}**\n${r.summary || 'No summary'}`)
                .join('\n\n');

        setMessages((prev) => [
          ...prev,
          { role: 'assistant', content: responseText, timestamp: new Date(), citations },
        ]);
      }
    } catch (err) {
      if (!abortRef.current) {
        setMessages((prev) => [
          ...prev,
          {
            role: 'assistant',
            content: `Error: ${err instanceof Error ? err.message : 'Unknown error'}`,
            timestamp: new Date(),
          },
        ]);
      }
    } finally {
      if (!abortRef.current) setIsGenerating(false);
    }
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
    <div className={styles.chatPanelContent}>
      <div className={styles.chatPanelHeader}>
        <span className={styles.chatPanelTitle}>
          {hasCopilot ? 'Copilot Chat' : 'Search Assistant'}
        </span>
        <button className={styles.chatPanelClear} onClick={handleClear}>
          Clear
        </button>
      </div>

      <div className={styles.chatPanelMessages}>
        {messages.length === 0 && (
          <div className={styles.chatPanelWelcome}>
            <Icon iconName="Robot" style={{ fontSize: 32, color: '#6264a7' }} />
            <p>Ask me anything about Bosch enterprise content.</p>
          </div>
        )}

        {messages.map((msg, i) => (
          <div
            key={i}
            className={`${styles.chatBubble} ${
              msg.role === 'user' ? styles.chatBubbleUser : styles.chatBubbleAssistant
            }`}
          >
            {msg.role === 'assistant' ? (
              <div dangerouslySetInnerHTML={{ __html: markdownToHtml(msg.content) }} />
            ) : (
              msg.content
            )}
            {msg.citations && msg.citations.length > 0 && (
              <div className={styles.chatBubbleCitations}>
                {msg.citations.map((cite, ci) => (
                  <a
                    key={ci}
                    href={cite.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className={styles.chatCitationLink}
                  >
                    [{ci + 1}] {cite.title}
                  </a>
                ))}
              </div>
            )}
          </div>
        ))}

        {isGenerating && !(messages.length > 0 && messages[messages.length - 1].role === 'assistant' && messages[messages.length - 1].content) && (
          <div className={styles.chatBubbleAssistant}>
            <Spinner label={hasCopilot ? 'Asking Copilot...' : 'Searching...'} />
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      <div className={styles.chatPanelInput}>
        <input
          type="text"
          placeholder={hasCopilot ? 'Ask Copilot...' : 'Ask a question...'}
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={isGenerating}
          className={styles.chatInputField}
        />
        <button
          className={styles.chatSendButton}
          onClick={() => handleSend()}
          disabled={isGenerating || !inputValue.trim()}
        >
          <Icon iconName="Send" />
        </button>
      </div>
    </div>
  );

  return (
    <>
      {!isOpen && (
        <button
          className={styles.chatFab}
          onClick={onOpen}
          title="Open Copilot Chat"
          aria-label="Open Copilot Chat"
        >
          <Icon iconName="Chat" style={{ fontSize: 24 }} />
        </button>
      )}

      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
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
