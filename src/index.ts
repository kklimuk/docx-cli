#!/usr/bin/env bun
import { main } from "./cli";

const code = await main(Bun.argv);
process.exit(code);
