# Contribution Guidelines

## Thank you for your interest in _xlsx_!

We're happy to see that you're interested in contributing to this
project. It is a labour of love and will continue to grow as more
people find use for it and give back.

## Document Utility

Following these guidelines helps to communicate that you respect the
time of the developers managing and developing this open source
project. In return, they should reciprocate that respect in addressing
your issue, assessing changes, and helping you finalize your pull
requests.

## How to Contribute

_xlsx_ is an open-source, MIT-licensed project and we appreciate
contributions from the community. Valid contributions include filing
bug reports, creating feature-requests, sponsoring development,
spreading the word, and so on. If you want to go the extra mile, we'll
also accept pull requests!

## Social Conduct

- Be nice, kind, and respectful.
- Don't go off-topic.

## Workflow

- Make sure that you have read the [README](../README.md).
- Create
  a [new issue](https://github.com/hungrybluedev/xlsx/issues/new/choose)
  and make either a Bug Report or a Feature Request.
- Fill in the fields to the best of your ability.
- Try to make minimal reproducible
  examples ([MREs](https://stackoverflow.com/help/minimal-reproducible-example))
  for bug reports. A small `.xlsx` file that triggers the problem is
  worth a great deal here.
- It is nice to provide as much context as possible so that developers
  have all the information they need to take care of your issue.
- If you feel that you can contribute code yourself, please _file an
  issue first_! This will help you figure out the work you need to do
  as well as help in keeping track of progress.

## Development Environment

### The V toolchain is pinned

The V commit this project builds against lives in [`.v-version`](../.v-version)
at the repository root. CI reads it through `vlang/setup-v`'s `version-file`
input, and the devcontainer builds that same commit from source, so local
builds and CI use the same compiler.

Please build with the pinned commit rather than whatever `v up` last gave you.

### Local setup

`spec/` and `examples/` consume this project the same way a user does, through
`import xlsx`, so the checkout has to be resolvable from `~/.vmodules`:

```sh
mkdir -p ~/.vmodules
ln -s "$(pwd)" ~/.vmodules/xlsx
```

Then the full set of gates CI runs:

```sh
v fmt -verify .                     # formatting gate
v check-md .                        # documentation gate
v run tools/check_readme_sync.v     # README examples match examples/snippets/
v should-compile-all examples/      # every example still compiles
v test .                            # module, spec/, and any examples/ tests
```

`v test .` walks the whole tree, so there is no separate command for the
`spec/` suite or for tests added under `examples/`.

### README examples are checked

`tools/check_readme_sync.v` requires the files in `examples/snippets/` to appear
verbatim in `README.md`. Edit the snippet file, then paste it into the README --
not the other way around, and don't reformat one copy only.

### Bumping the pinned V commit

Dependabot cannot do this for us. It has no V ecosystem, and for GitHub Actions
it only rewrites the `uses:` reference of an action, never the `with:` inputs --
so the value in `.v-version` is invisible to it. Pinning to a V tag is not a
workaround either: V's tags lag master by months, and its `0.5.2` release tag is
re-published in place rather than being immutable.

Instead, the **V master canary** job in `ci.yml` builds against V master on every
push and once a week. It is never a merge gate. When it goes red:

1. Investigate whether the breakage is ours or an upstream V regression.
2. If it is ours, fix it and bump `.v-version` to the master commit you tested.
3. If it is upstream, leave `.v-version` alone and wait.

Bump the pin periodically even when nothing is broken. V's `GNUmakefile` clones
its bootstrap compiler (`vlang/vc`) unpinned, so a very old commit eventually
stops building from source at all.

### Project structure

V's layout rules changed, and two of them are easy to trip over:

- **There is no `src/` directory.** V removed support for treating `src/` as a
  virtual module root and now errors out if it finds sources there. Every file
  of the `xlsx` module lives at the repository root, next to `v.mod`. If you ever
  do want a source subfolder, the supported mechanism is the explicit `base_url`
  field in `v.mod` -- not a bare `src/`.
- **`subdirs:` in `v.mod` is for splitting one module across directories**, where
  every file still declares the *same* module. It is not how you register
  submodules, and it is not for test fixtures. `spec/` and `examples/` are
  standalone programs that `import xlsx`, so they must *not* be listed there.

Tests that write files should write them under `os.temp_dir()` and clean up in
`testsuite_end`, so a test run leaves the working tree clean.
