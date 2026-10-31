# Releasing

Releases are published from a tag by `.github/workflows/release.yml` through npm **trusted publishing** (OIDC: no npm tokens, no secrets) and **staged publishing**: the workflow only stages the packages, and a maintainer approves them with two-factor authentication before they become public. Provenance is generated automatically by trusted publishing. Sources of these facts are in docs/decisions.md (F-NPM-01 … F-NPM-03).

Published packages: `xlvane` and `@xlvane/exceljs`, always with the same version.

## Once, before the first release

1. **npm account** with two-factor authentication (a security key or an authenticator app).
2. **npm organization `xlvane`** (free for public packages), so that `@xlvane/exceljs` can be published.
3. **Public repository**: provenance is not generated for private repositories.
4. **GitHub environment `npm-release`** (Settings → Environments). Optional: add yourself as a required reviewer, so that the publish job waits for your approval.
5. **Private vulnerability reporting** (Settings → Security → Private vulnerability reporting), which SECURITY.md refers to.

## First release (1.0.0)

npm configures a trusted publisher in the settings of an existing package, so the first release creates the packages first:

1. Merge to `main` and check that CI is green on all jobs.
2. Tag and push: `git tag v1.0.0 && git push origin v1.0.0`. The `verify` job runs all tests; the `publish` job is expected to fail at this point (no trusted publisher yet).
3. Create the packages from your machine (needs npm 11.15.0 or later: `npm install -g npm@11.21.1`), in a clean checkout of the tag:
   ```sh
   npm login
   npm ci && npm run build
   npm stage publish -w packages/xlvane --provenance=false
   npm stage publish -w packages/exceljs --provenance=false
   ```
   `--provenance=false` is needed because `publishConfig.provenance` is `true` and provenance can only be generated in a supported CI; a command-line flag takes precedence over `publishConfig` (npm 11.21.1 lib/commands/publish.js). npm may print `"bin[xlvane]" script name dist/cli.js was invalid and removed`: it only normalizes `./dist/cli.js` to `dist/cli.js` and keeps the entry (@npmcli/package-json lib/normalize.js). For a package that does not exist yet, npm publishes a public placeholder version `0.0.0-stage`; the staged 1.0.0 stays private until approved.
4. Reject these local stages, because they have no provenance: `npm stage list`, then `npm stage reject <stage-id>` for each.
5. On npmjs.com, for **each** package: Settings → Trusted publishing → GitHub Actions: organization or user `DevXord`, repository `xlvane`, workflow filename `release.yml`, environment `npm-release`; allowed action: `npm stage publish` only. The configuration must be used within 2 days, or it expires.
6. In GitHub Actions, re-run the failed `publish` job of the `v1.0.0` run.
7. Review and approve:
   ```sh
   npm stage list
   npm stage download <stage-id>   # inspect the tarball
   npm stage approve <stage-id>    # asks for 2FA
   ```
   (or the Staged Packages tab on npmjs.com).
8. On npmjs.com, for each package: Settings → Publishing access → "Require two-factor authentication and disallow tokens".

If npm refuses to stage 1.0.0 again after step 4 (the npm documentation does not say whether a rejected staged version can be staged again), skip the rejection in step 4 and approve the CI stage only, or ask npm support; do not approve a stage that has no provenance.

## Every later release

1. Bump the version of both packages in one commit (`packages/xlvane/package.json`, `packages/exceljs/package.json` and its `xlvane` dependency `^X.Y.Z`), update CHANGELOG.md.
2. Tag `vX.Y.Z` and push the tag. `scripts/check-release-version.mjs` stops the release if the tag and the versions differ.
3. When the workflow has staged both packages: `npm stage list`, inspect, `npm stage approve <stage-id>` (2FA).

## After a release

- `npm view xlvane dependencies` lists only `fflate`.
- The package page on npmjs.com shows the provenance of the release.
- `npm audit signatures` in a project that installs the packages verifies their signatures and attestations.
