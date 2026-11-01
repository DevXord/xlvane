# Releasing

Releases are published from a tag by `.github/workflows/release.yml` through npm **trusted publishing** (OIDC: no npm tokens, no secrets) and **staged publishing**: the workflow only stages the packages, and a maintainer approves them with two-factor authentication before they become public. Provenance is generated automatically by trusted publishing. Sources of these facts are in docs/decisions.md (F-NPM-01 … F-NPM-03).

Published packages: `xlvane` and `@xlvane/compat`, always with the same version.

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

## Every later release

1. Bump the version of both packages in one commit (`packages/xlvane/package.json`, `packages/compat/package.json` and its `xlvane` dependency `^X.Y.Z`), update CHANGELOG.md.
2. Tag `vX.Y.Z` and push the tag. `scripts/check-release-version.mjs` stops the release if the tag and the versions differ.
3. When the workflow has staged both packages: `npm stage list`, inspect, `npm stage approve <stage-id>` (2FA).

## After a release

- `npm view xlvane dependencies` lists only `fflate`.
- The package page on npmjs.com shows the provenance of the release.
- `npm audit signatures` in a project that installs the packages verifies their signatures and attestations.
