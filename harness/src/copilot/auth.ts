import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import { execSync } from 'node:child_process';

export interface CopilotTokenInfo {
  token: string;
  expiresAt: number;
  apiEndpoint: string;
}

export class CopilotAuthManager {
  private cachedToken: CopilotTokenInfo | null = null;
  private explicitGitHubToken?: string;

  constructor(gitHubToken?: string) {
    this.explicitGitHubToken = gitHubToken;
  }

  /**
   * Retrieves a valid GitHub Copilot access token, refreshing if necessary.
   */
  async getCopilotToken(): Promise<CopilotTokenInfo> {
    const now = Math.floor(Date.now() / 1000);
    // If cached and still valid for at least 5 minutes, reuse
    if (this.cachedToken && this.cachedToken.expiresAt > now + 300) {
      return this.cachedToken;
    }

    // Direct Copilot token passed via environment
    if (process.env.COPILOT_API_TOKEN || process.env.COPILOT_TOKEN) {
      const token = (process.env.COPILOT_API_TOKEN || process.env.COPILOT_TOKEN)!;
      this.cachedToken = {
        token,
        expiresAt: now + 86400, // 24 hours default
        apiEndpoint: process.env.COPILOT_API_URL || 'https://api.githubcopilot.com',
      };
      return this.cachedToken;
    }

    const ghToken = this.resolveGitHubToken();
    if (!ghToken) {
      throw new Error(
        'No GitHub token found. Please set GITHUB_TOKEN / GH_TOKEN environment variable or run `gh auth login`.'
      );
    }

    const tokenInfo = await this.exchangeGitHubTokenForCopilot(ghToken);
    this.cachedToken = tokenInfo;
    return tokenInfo;
  }

  /**
   * Resolve a GitHub OAuth/personal token from env, configs, or gh CLI.
   */
  private resolveGitHubToken(): string | null {
    if (this.explicitGitHubToken) {
      return this.explicitGitHubToken;
    }

    if (process.env.GH_TOKEN) {
      return process.env.GH_TOKEN;
    }
    if (process.env.GITHUB_TOKEN) {
      return process.env.GITHUB_TOKEN;
    }

    // Check ~/.config/github-copilot/hosts.json
    const copilotHostsPath = path.join(os.homedir(), '.config', 'github-copilot', 'hosts.json');
    if (fs.existsSync(copilotHostsPath)) {
      try {
        const data = JSON.parse(fs.readFileSync(copilotHostsPath, 'utf8'));
        for (const host of Object.keys(data)) {
          if (data[host]?.oauth_token) {
            return data[host].oauth_token;
          }
        }
      } catch {
        // Ignore read/parse error
      }
    }

    // Check ~/.config/github-copilot/apps.json
    const copilotAppsPath = path.join(os.homedir(), '.config', 'github-copilot', 'apps.json');
    if (fs.existsSync(copilotAppsPath)) {
      try {
        const data = JSON.parse(fs.readFileSync(copilotAppsPath, 'utf8'));
        for (const host of Object.keys(data)) {
          if (data[host]?.oauth_token) {
            return data[host].oauth_token;
          }
        }
      } catch {
        // Ignore read/parse error
      }
    }

    // Check ~/.config/gh/hosts.yml
    const ghHostsPath = path.join(os.homedir(), '.config', 'gh', 'hosts.yml');
    if (fs.existsSync(ghHostsPath)) {
      try {
        const content = fs.readFileSync(ghHostsPath, 'utf8');
        const match = content.match(/oauth_token:\s*([^\s\r\n]+)/);
        if (match && match[1]) {
          return match[1];
        }
      } catch {
        // Ignore read error
      }
    }

    // Try `gh auth token` as subprocess fallback
    try {
      const output = execSync('gh auth token', {
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'ignore'],
      }).trim();
      if (output && output.length > 5) {
        return output;
      }
    } catch {
      // gh cli might not be authenticated
    }

    return null;
  }

  /**
   * Exchanges a GitHub token for an internal Copilot token.
   */
  private async exchangeGitHubTokenForCopilot(ghToken: string): Promise<CopilotTokenInfo> {
    const url = 'https://api.github.com/copilot_internal/v2/token';
    const response = await fetch(url, {
      method: 'GET',
      headers: {
        Authorization: `token ${ghToken}`,
        Accept: 'application/json',
        'User-Agent': 'GitHubCopilotChat/0.24.0',
        'Editor-Version': 'vscode/1.95.0',
        'Editor-Plugin-Version': 'copilot-chat/0.24.0',
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Failed to exchange GitHub token for Copilot token (${response.status}): ${errorText}`
      );
    }

    const data = (await response.json()) as {
      token: string;
      expires_at: number;
      endpoints?: { api?: string };
    };

    return {
      token: data.token,
      expiresAt: data.expires_at,
      apiEndpoint: data.endpoints?.api || 'https://api.githubcopilot.com',
    };
  }
}
