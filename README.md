# Teams Office

Teams Office is [Teams for Linux](https://github.com/IsmaelMartinez/teams-for-linux)
with a local MCP bridge for AI agents. Agents can open Word and Excel documents
in background tabs, create Word documents in existing SharePoint folders, and
read and edit documents through an Office add-in. Office handles saving and
coauthoring as usual.

This is an unofficial client, independent of Microsoft and the upstream Teams
for Linux project.

## Install

Install the [Teams Office Flatpak](https://github.com/nosini/teams-office-flatpak):

```sh
flatpak remote-add --user --if-not-exists flathub https://dl.flathub.org/repo/flathub.flatpakrepo
flatpak install --user https://nosini.github.io/teams-office-flatpak/teams-office.flatpakref
```

Start **Teams Office** from your application menu and sign in to Teams. It
installs alongside the regular Teams for Linux client with its own settings.

## Connect your agent

The Flatpak enables the bridge automatically. Point your agent's Teams Office
MCP connection at:

```text
~/.var/app/eu.nosini.TeamsOffice/config/teams-for-linux/mcp/bridge.sock
```

Open a Word or Excel document, then choose **Documents > Connect Bridge**, or
ask your agent to set up the bridge. Sign-in, MFA and your organization's
Office add-in policy can require your interaction.

The [agent guide](docs/office-agents.md) explains connection setup, exporting a
standalone relay, container connections and the available tools.

## What agents can do

- Open Word and Excel documents in independent background tabs.
- Find existing SharePoint folders and create Word documents there.
- Read and edit Excel values, formulas, formatting and worksheet structure.
- Read Word paragraphs and tables, edit text and styles, work with comments,
  and review tracked changes.

Reads return short-lived tokens. If the content changes before an edit, the
agent must read it again. Microsoft permissions still apply. Experimental
Markdown editing and channel-file tools require separate configuration.

## Local access

The bridge uses a private Unix socket. The Office add-in is served over HTTPS
on loopback with a certificate trusted only by this client. Expose only the
socket to an agent container: the adjacent add-in directory contains the
certificate's private key. Connected agents can access documents through the
enabled tools; give the socket only to agents you trust.

## License and development

Teams for Linux and the Office bridge are licensed under
[GPL-3.0-or-later](LICENSE.md). Upstream retains its own source, history and
documentation under `upstream/teams-for-linux`.

See [development](docs/development.md) for the source layout, building, testing
and updating the pinned upstream version.
