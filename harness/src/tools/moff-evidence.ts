import {
  defineDoc, defineExtension, hook, ToolTask,
  type ToolExecutionApi,
} from '@earendil-works/pi-durable';
import type { Context } from '@earendil-works/chord';

const Evidence = defineDoc({
  kind: 'moff.read-evidence', version: 1, scope: 'conversation', history: 'latest', fork: 'initial',
  initial: (): { ids: string[] } => ({ ids: [] }),
});

// Record opaque identifiers from actual JSON results, without copying mailbox bodies.
export async function recordMoffEvidence(output: string, api: ToolExecutionApi, context: Context): Promise<void> {
  if (!/^\s*[\[{]/.test(output)) return; // CLI --help is plain text, not record data.
  const value: unknown = JSON.parse(output);
  const ids = new Set<string>();
  const visit = (node: unknown) => {
    if (!node || typeof node !== 'object') return;
    for (const [key, child] of Object.entries(node)) {
      if (/id$/i.test(key) && typeof child === 'string' && child) ids.add(child);
      else visit(child);
    }
  };
  visit(value);
  if (!ids.size) return;
  await api.commit(async tx => {
    const document = await tx.doc(Evidence, api.conversationId);
    document.ids = [...new Set([...document.ids, ...ids])];
  }, context);
}

export const MoffEvidenceGuard = defineExtension({
  name: 'moff-evidence-guard',
  hooks: [hook(ToolTask, {
    beforeTool: async (call, api, context) => {
      if (call.name !== 'moff') return;
      const args = call.arguments.args;
      if (!Array.isArray(args) || args.includes('--help')) return;
      const evidence = await api.snapshot(Evidence, api.conversationId, context);
      for (let index = 0; index < args.length; index++) {
        const flag = args[index];
        if (typeof flag !== 'string') continue;
        const match = /^--(?:[a-z-]+-)?id(?:=(.*))?$/.exec(flag);
        if (!match) continue;
        const value = match[1] ?? args[index + 1];
        if (typeof value === 'string' && !evidence?.ids.includes(value)) {
          return { block: `The ${flag.split('=')[0]} value was not discovered in a Moff result. ` +
            'Do not invent IDs such as primary. Discover the actual record/folder with Moff read commands first; ' +
            'calendar folders is ["outlook","calendar","folders"]. Use --help if needed.' };
        }
      }
    },
  })],
});
