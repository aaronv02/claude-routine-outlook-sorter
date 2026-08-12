/**
 * A mailbox in memory, served over a stubbed `fetch`.
 *
 * This exists so the real commands can be run end to end - the actual `plan` and
 * `apply` from `routine/sweep.ts`, the actual `src/graph.ts` with its JSON
 * batching and provenance encoding, the actual promotion into native rules -
 * without a Microsoft account.
 *
 * It intercepts at `fetch` rather than mocking the Graph module, which is the whole
 * point: a mock of `graph.ts` would skip exactly the parts most likely to be wrong.
 * URL construction, `$batch` sub-response handling, the provenance round-trip, and
 * the fact that a PATCH on `categories` replaces the array are all live here.
 */

export interface FakeMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  receivedDateTime: string;
  hasAttachments: boolean;
  categories: string[];
  from: { emailAddress: { address: string; name: string } };
  /** Our provenance stamp, as it would sit on the real message. */
  stamp?: string;
  /** Present on the messages used to seed the first run, which live outside the inbox. */
  filed?: boolean;
}

export interface FakeRule {
  id: string;
  displayName: string;
  sequence?: number;
  isEnabled?: boolean;
  conditions?: { senderContains?: string[] };
  actions?: { assignCategories?: string[]; stopProcessingRules?: boolean };
}

const PROVENANCE_PROP_ID = "String {c11ff724-aa03-4555-9952-8fa248a11c3e} Name InboxStewardAssignment";

export class FakeMailbox {
  messages: FakeMessage[] = [];
  masterCategories: { id: string; displayName: string; color: string }[] = [];
  rules: FakeRule[] = [];
  /** The hidden state folder, once created. */
  private stateFolderId: string | null = null;
  private stateMessage: { id: string; content: string } | null = null;

  /** Every request path seen, so a test can assert what was and wasn't called. */
  readonly calls: string[] = [];

  private nextId = 1;

  add(message: Omit<FakeMessage, 'id'> & { id?: string }): FakeMessage {
    const created = { ...message, id: message.id ?? `msg-${this.nextId++}` } as FakeMessage;
    this.messages.push(created);
    return created;
  }

  /** What Outlook would show for a message, for assertions. */
  categoriesOn(id: string): string[] {
    return this.messages.find((m) => m.id === id)?.categories ?? [];
  }

  /**
   * Her changing a category by hand in Outlook.
   *
   * The stamp is deliberately left alone - that is exactly what happens in the real
   * mailbox, and it is the divergence between stamp and categories that the next
   * run reads as a correction.
   */
  sheRecategorizes(id: string, categoryNames: string[]): void {
    const message = this.messages.find((m) => m.id === id);
    if (!message) throw new Error(`no such message: ${id}`);
    message.categories = categoryNames;
  }

  install(): () => void {
    const original = globalThis.fetch;
    globalThis.fetch = ((input: string | URL | Request, init?: RequestInit) =>
      this.handle(String(input), init)) as typeof fetch;
    return () => {
      globalThis.fetch = original;
    };
  }

  // -------------------------------------------------------------------------

  private json(body: unknown, status = 200): Response {
    return new Response(JSON.stringify(body), {
      status,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  private async handle(url: string, init?: RequestInit): Promise<Response> {
    const method = (init?.method ?? 'GET').toUpperCase();
    const path = url.replace('https://graph.microsoft.com/v1.0', '');
    this.calls.push(`${method} ${path.split('?')[0]}`);

    // --- Token endpoint --------------------------------------------------
    if (url.includes('login.microsoftonline.com')) {
      return this.json({
        access_token: 'fake-access-token',
        refresh_token: 'fake-refresh-token-rotated',
        expires_in: 3600,
        scope:
          'Mail.ReadWrite MailboxSettings.ReadWrite Calendars.Read User.Read offline_access',
      });
    }

    // --- Whoami ----------------------------------------------------------
    if (path.startsWith('/me?')) {
      return this.json({ displayName: 'The Director', mail: 'director@example.org' });
    }

    // --- Master category list --------------------------------------------
    if (path.startsWith('/me/outlook/masterCategories')) {
      if (method === 'POST') {
        const body = JSON.parse(String(init?.body ?? '{}')) as {
          displayName: string;
          color: string;
        };
        const created = { id: `cat-${this.nextId++}`, ...body };
        this.masterCategories.push(created);
        return this.json(created, 201);
      }
      return this.json({ value: this.masterCategories });
    }

    // --- The hidden state folder -----------------------------------------
    if (path.startsWith('/me/mailFolders?')) {
      const folders = this.stateFolderId
        ? [{ id: this.stateFolderId, displayName: 'Inbox Steward' }]
        : [];
      return this.json({ value: folders });
    }
    if (path === '/me/mailFolders' && method === 'POST') {
      this.stateFolderId = 'folder-state';
      return this.json({ id: this.stateFolderId }, 201);
    }
    if (this.stateFolderId && path.startsWith(`/me/mailFolders/${this.stateFolderId}/messages`)) {
      if (method === 'POST') {
        const body = JSON.parse(String(init?.body ?? '{}')) as { body: { content: string } };
        this.stateMessage = { id: 'state-msg', content: body.body.content };
        return this.json({ id: 'state-msg' }, 201);
      }
      return this.json({
        value: this.stateMessage
          ? [{ id: this.stateMessage.id, body: { content: this.stateMessage.content } }]
          : [],
      });
    }
    if (path === '/me/messages/state-msg' && method === 'PATCH') {
      const body = JSON.parse(String(init?.body ?? '{}')) as { body: { content: string } };
      if (this.stateMessage) this.stateMessage.content = body.body.content;
      return this.json({ id: 'state-msg' });
    }

    // --- Native rules ----------------------------------------------------
    if (path.startsWith('/me/mailFolders/inbox/messageRules')) {
      if (method === 'POST') {
        const body = JSON.parse(String(init?.body ?? '{}')) as FakeRule;
        const created = { ...body, id: `rule-${this.nextId++}` };
        this.rules.push(created);
        return this.json(created, 201);
      }
      if (method === 'PATCH') {
        const id = path.split('/').pop() as string;
        const body = JSON.parse(String(init?.body ?? '{}')) as Partial<FakeRule>;
        const existing = this.rules.find((r) => r.id === id);
        if (existing) Object.assign(existing, body);
        return new Response(null, { status: 204 });
      }
      return this.json({ value: this.rules });
    }

    // --- Reading mail ----------------------------------------------------
    if (path.startsWith('/me/mailFolders/inbox/messages')) {
      return this.json({ value: this.messages.filter((m) => !m.filed).map((m) => this.render(m)) });
    }
    if (path.startsWith('/me/messages?')) {
      // listCategorizedHistory: everything, anywhere in the mailbox.
      return this.json({ value: this.messages.map((m) => this.render(m)) });
    }

    // --- Batched writes --------------------------------------------------
    if (path === '/$batch' && method === 'POST') {
      const body = JSON.parse(String(init?.body ?? '{}')) as {
        requests: { id: string; url: string; body: Record<string, unknown> }[];
      };

      const responses = body.requests.map((request) => {
        const id = request.url.replace('/me/messages/', '');
        const message = this.messages.find((m) => m.id === id);
        if (!message) return { id: request.id, status: 404, body: { error: 'not found' } };

        // The real PATCH replaces the array wholesale. Reproducing that is the
        // point: it is what makes losing one of her own categories possible.
        if (Array.isArray(request.body.categories)) {
          message.categories = request.body.categories as string[];
        }
        const extended = request.body.singleValueExtendedProperties as
          | { id: string; value: string }[]
          | undefined;
        const stamp = extended?.find((p) => p.id === PROVENANCE_PROP_ID);
        if (stamp) message.stamp = stamp.value;

        return { id: request.id, status: 200, body: {} };
      });

      return this.json({ responses });
    }

    return this.json({ error: `fake mailbox has no route for ${method} ${path}` }, 404);
  }

  private render(m: FakeMessage): Record<string, unknown> {
    return {
      id: m.id,
      subject: m.subject,
      bodyPreview: m.bodyPreview,
      receivedDateTime: m.receivedDateTime,
      hasAttachments: m.hasAttachments,
      categories: m.categories,
      from: m.from,
      ...(m.stamp
        ? { singleValueExtendedProperties: [{ id: PROVENANCE_PROP_ID, value: m.stamp }] }
        : {}),
    };
  }
}
