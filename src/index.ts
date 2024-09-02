#!/usr/bin/env node

import { cli } from './cli/cli';

// required to make console.log() in combination with piped output synchronous
// on Windows/in PowerShell so that the output is not trimmed by calling
// process.exit() after executing the command, while the output is still
// being processed; https://github.com/pnp/cli-microsoft365/issues/1266
if ((process.stdout as any)._handle) {
  (process.stdout as any)._handle.setBlocking(true);
}

try {
  cli.execute(process.argv.slice(2));
}
catch (e: any) {
  process.exit(1);
}
