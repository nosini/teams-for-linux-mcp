# Development

This repository owns the Office bridge, its documentation and build
integration. Teams for Linux is an unmodified Git submodule pinned to a
specific commit. Upstream updates do not replace this repository's README or
workflows.

## Source layout

```text
src/                       Office, MCP and editor source added to app/
tools/                     MCP command-line helpers added to scripts/
tests/                     Office bridge unit, fixture and Electron tests
test/                      Build assembly tests
docs/                      Our user and developer documentation
patches/                   Integration changes to upstream files
dependencies/package-lock.json  Lockfile for the assembled app
package.overrides.json     Our additions to the upstream package manifest
upstream/teams-for-linux/  Pinned upstream submodule
build/teams-office/        Generated app, dependencies and build output
```

`scripts/prepare-app.js` copies committed upstream files into a temporary build
directory, applies patches in filename order, adds our files and merges the
package overrides. It checks the dependency lock before replacing the
previous assembly. Installed dependencies and build artifacts are retained.
Deleted source files disappear from the next assembly. The script also works
when flatpak-builder has removed Git metadata from a source checkout.

Files under `src/` become files under the assembled app's `app/` directory.
Their existing relative imports and the upstream security boundaries stay in
place. Our tests run alongside upstream's tests in the assembled directory.
The four integration points are app startup, configuration, the Documents menu
entry and the Documents shell's allowed IPC channels.

## Set up

Use Node.js 24 and npm. Clone with the submodule:

```sh
git clone --recurse-submodules https://github.com/nosini/teams-for-linux-mcp.git
cd teams-for-linux-mcp
npm run install:app
```

For an existing checkout, run `git submodule update --init --recursive` first.
The root package has no dependencies; the app's dependencies live under
`build/teams-office/node_modules`. Install scripts are skipped so dependency
installation does not download Electron or Playwright browsers implicitly.

To run the desktop client, install Electron's pinned binary explicitly:

```sh
cd build/teams-office
node node_modules/electron/install.js
cd ../..
npm start -- --mcp.enabled --mcp.addin.enabled
```

Electron needs the usual Linux desktop libraries and OpenSSL for the add-in
certificate. In an environment where Electron's bundled ZIP extractor leaves
an incomplete `dist/`, extract the downloaded Electron ZIP with `unzip` into
`build/teams-office/node_modules/electron/dist/` and write `electron` into
`build/teams-office/node_modules/electron/path.txt`.

The app keeps its existing package name and profile paths. This source
reorganization does not migrate or reset sign-in data.

## Test and build

```sh
npm test
npm run lint
npm run pack
```

`npm test` checks build assembly and runs the complete upstream and Office
unit suite. `npm run lint` checks both our source and the assembled app.
`npm run pack` uses upstream's electron-builder configuration with our package
overrides and disables publishing. Its output is under
`build/teams-office/dist/`.

Commands such as `npm run mcp:diagnose`, `npm run mcp:call -- TOOL JSON` and
`npm run mcp:setup -- /absolute/export-directory` prepare the app and run its
corresponding helper. They require the app dependencies to be installed.

Electron fixtures are under `tests/electron/` and run from the assembly, for
example:

```sh
cd build/teams-office
xvfb-run -a node_modules/.bin/electron --no-sandbox tests/electron/documentCreationSmoke.js
```

These fixtures simulate Microsoft services and Office.js. They check Electron
integration without accessing real accounts; live Microsoft behavior needs a
separate acceptance test. See [Office bridge internals](development/office-mcp.md)
for the remaining fixtures and the protocol details. Paths beginning with
`app/` in that guide refer to the assembled directory.

Our fixtures keep their temporary profiles under the repository's
`build/t/` directory. For a deeply nested checkout, set
`TEAMS_MCP_TEST_TMP` to a shorter project scratch directory so Unix socket
paths fit within the platform's length limit.

## Update upstream

Choose a specific upstream release or commit, then:

```sh
git -C upstream/teams-for-linux fetch origin
git -C upstream/teams-for-linux checkout --detach COMMIT_OR_TAG
npm run prepare:app
npm run install:app
npm test
npm run lint
```

The assembly fails when a patch no longer applies or upstream introduces a
file at one of our source paths. Review the integration and adjust the patch
or our file before continuing. Changes belong in `src/`, `tools/`, `tests/` or
`patches/`; edits made directly in the generated app are replaced on prepare.

If upstream's lockfile or our dependency requirements change, `prepare:app` fails
with a lockfile error. Run `npm run update:lock`, review the resulting
`dependencies/package-lock.json`, then install and test again. Updating the
lockfile starts from upstream's dependency resolutions, including transitive
fixes, then adds our requirements. It can use the network; normal assembly is
offline. `dependencies/upstream-lock.sha256` records which upstream lock was
used. Record it with the submodule pointer, integration changes and app lock.

To edit an integration patch, prepare the app and make the intended edit in
`build/teams-office/app/`. Compare the affected files against the pinned
submodule and replace the relevant patch, then prepare again to verify it.
Never commit modifications inside the upstream submodule.

## Flatpak releases

[teams-office-flatpak](https://github.com/nosini/teams-office-flatpak) owns
Flatpak packaging and publication. Its app source must check out this
repository with submodules, run `node scripts/prepare-app.js`, and build from
`build/teams-office`. Generate its npm cache from
`dependencies/package-lock.json`, which is the assembled app's complete lock.

Source releases retain the `office-vVERSION` tag convention, with `-N` for
additional releases on the same upstream app version. Pin a tested source tag
and commit in the Flatpak manifest. Existing tags retain their original trees;
new-layout build commands need a new source tag containing this layout.
