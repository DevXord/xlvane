# Releasing

Releases are published from a tag by `.github/workflows/release.yml` through npm **trusted publishing** (OIDC: no npm tokens, no secrets) and **staged publishing**: the workflow only stages the packages, and a maintainer approves them with two-factor authentication before they become public. Provenance is generated automatically by trusted publishing. Sources of these facts are in docs/decisions.md (F-NPM-01 … F-NPM-03).

Published packages: `xlvane`, `@xlvane/compat` and (from 1.2.0) `@xlvane/mcp`, always with the same version.

## Once, before the first release

1. **npm account** with two-factor authentication (a security key or an authenticator app).
2. **npm organization `xlvane`** (free for public packages), so that `@xlvane/compat` can be published.
3. **Public repository**: provenance is not generated for private repositories.
4. **GitHub environment `npm-release`** (Settings → Environments). Optional: add yourself as a required reviewer, so that the publish job waits for your approval.
5. **Private vulnerability reporting** (Settings → Security → Private vulnerability reporting), which SECURITY.md refers to.

## First release (1.0.0)

npm configures a trusted publisher in the settings of an existing package, so the first release is published from the maintainer's machine, without provenance (D-R10); every later release goes through the workflow with provenance.

1. Merge to `main` and check that CI is green on all jobs.
2. Tag and push: `git tag v1.0.0 && git push origin v1.0.0`. The `verify` job runs all tests. The `publish` job waits for approval in the `npm-release` environment: **reject it** (Review deployments → Reject), because the packages do not exist on npm yet.
3. Publish from your machine (needs npm 11.15.0 or later: `npm install -g npm@11.21.1`), in a clean checkout of the tag:
   ```sh
   git switch --detach v1.0.0
   npm login
   npm ci && npm run build
   npm stage publish -w packages/xlvane --provenance=false
   npm stage publish -w packages/compat --provenance=false
   npm stage list
   npm stage download <stage-id>   # optional: inspect the tarball
   npm stage approve <stage-id>    # for each package; asks for 2FA
   ```
   - `--provenance=false` is needed because `publishConfig.provenance` is `true` and provenance can only be generated in a supported CI; a command-line flag takes precedence over `publishConfig` (npm 11.21.1 lib/commands/publish.js).
   - npm may print `"bin[xlvane]" script name dist/cli.js was invalid and removed`: it only normalizes `./dist/cli.js` to `dist/cli.js` and keeps the entry (@npmcli/package-json lib/normalize.js).
   - For a package that does not exist yet, npm publishes a public placeholder version `0.0.0-stage`; the staged 1.0.0 stays private until approved.
   - **Do not run `npm stage reject` on the first stage of a new package.** On 2026-10-31 rejecting both first stages removed the packages entirely (`npm view xlvane`: "Unpublished on 2026-10-31T19:35:12.446Z"), and npm's unpublish policy then blocks the name: "If you unpublish an entire package, you may not publish any new versions of that package until 24 hours have passed"; a `name@version` that was published once cannot be reused (npm/documentation `unpublishing-packages-from-the-registry.mdx`, read 2026-10-31).
4. Check: `npm view xlvane version` and `npm view @xlvane/compat version` print `1.0.0`; `npm view xlvane dependencies` lists only `fflate`.
5. On npmjs.com, for **each** package: Settings → Trusted publishing → GitHub Actions: organization or user `DevXord`, repository `xlvane`, workflow filename `release.yml`, environment `npm-release`; allowed action: `npm stage publish` only. The configuration must be used within 2 days, or it expires (F-NPM-01); if it expires before the next release, add it again then.
6. On npmjs.com, for each package: Settings → Publishing access → "Require two-factor authentication and disallow tokens".

## First release of `@xlvane/mcp` (1.2.0)

npm sets up a trusted publisher only for a package that exists (F-NPM-05), so `@xlvane/mcp` 1.2.0 is published from your machine, like `xlvane` 1.0.0 (D-R10, D-I03). The workflow stages `xlvane` and `@xlvane/compat` 1.2.0 as usual and only prints a notice for `@xlvane/mcp`.

1. Steps 1 and 3 to 5 of "Every later release" below for `xlvane` and `@xlvane/compat`; approve their stages first.
2. In a clean checkout of the tag (npm 11.15.0 or later):
   ```sh
   git switch --detach v1.2.0
   npm login
   npm ci && npm run build
   npm stage publish -w packages/mcp --provenance=false
   npm stage list
   npm stage approve <stage-id>    # asks for 2FA; never reject the first stage of a new package
   ```
3. Check: `npm view @xlvane/mcp version` prints `1.2.0`, and `npm view @xlvane/mcp dependencies` lists `@modelcontextprotocol/server` and `xlvane` only.
4. On npmjs.com for `@xlvane/mcp`: Settings → Trusted publishing → GitHub Actions with the same values as the other packages (step 5 of "First release"), and Settings → Publishing access → "Require two-factor authentication and disallow tokens". The configuration must be used by the next release within 2 days, or it expires (F-NPM-01); if it expires, add it again before that release.

## Every later release

1. Bump the version of all packages in one commit (`packages/xlvane/package.json`, `packages/compat/package.json` and `packages/mcp/package.json` with their `xlvane` dependency `^X.Y.Z`, `SERVER_VERSION` in `packages/mcp/src/server.ts`, and both `version` fields of `packages/mcp/server.json`), update CHANGELOG.md.
   Right before tagging, replace the "unreleased" markers of the release in one commit: the CHANGELOG heading gets the day of tagging (UTC; the publication day in the registry can follow it), and the README comparison, ROADMAP.md and docs/facts/comparison.md say the version is released. The tag freezes these files, and the website builds its changelog page from the tag (D-K03).
2. Check the trusted publisher of each package: `npm trust list xlvane`, `npm trust list @xlvane/compat` and `npm trust list @xlvane/mcp` print `github`, `release.yml`, `DevXord/xlvane`, `npm-release` and `stage publish` (F-NPM-04). A configuration that was added more than 2 days ago and never used has expired: delete it and add it again (F-NPM-01).
3. Tag `vX.Y.Z` and push the tag. `scripts/check-release-version.mjs` stops the release if the tag and the versions differ.
4. When the `verify` job has passed, approve the `publish` job (Review deployments → `npm-release` → Approve and deploy). `scripts/stage-publish.mjs` stages each package whose version is not on npm yet; if one fails (for example E401 when its trusted publisher is missing), the others are still staged and the job fails at the end. After fixing the cause, re-run the job: versions already published are skipped (a version that is staged but not approved yet is tried again, which npm refuses; approve or let it expire first).
5. When the workflow has staged the packages: `npm stage list`, inspect, `npm stage approve <stage-id>` (2FA). Approve `xlvane` first: the other packages depend on its new version.

## MCP Registry (after each release, from 1.2.1)

The official MCP Registry stores metadata only; it lists `@xlvane/mcp` as `io.github.DevXord/xlvane` (with this case: the registry compares the namespace with case, D-I07) from `packages/mcp/server.json` and checks that `mcpName` in the published `package.json` matches (D-I06, F-MCP-03). When the new `@xlvane/mcp` version is on npm:

1. Install `mcp-publisher` once, from the registry's releases page or with `brew install mcp-publisher`.
2. Log in:
   ```sh
   mcp-publisher login github
   ```
   It opens GitHub's device flow; the personal namespace `io.github.<your-user>/` is granted. The login expires after some time; repeat it when `publish` says the token is invalid or expired.
3. In a checkout of the release tag:
   ```sh
   cd packages/mcp
   mcp-publisher validate
   mcp-publisher publish
   ```
4. Check:
   ```sh
   curl "https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.DevXord/xlvane"
   ```
   It lists the new version.

## After a release

- `npm view xlvane dependencies` lists only `fflate`; `npm view @xlvane/mcp dependencies` lists only `@modelcontextprotocol/server` and `xlvane`.
- The package page on npmjs.com shows the provenance of the release.
- `npm audit signatures` in a project that installs the packages verifies their signatures and attestations.
- Rebuild the website in the repository `xlvane/xlvane.github.io` (D-G01): `node scripts/build.mjs --version X.Y.Z`, then `node scripts/check.mjs` and `node --test test/*.test.mjs`; commit `docs/` and push. The pages are generated from the packages just published on npm.
