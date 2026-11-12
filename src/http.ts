#!/usr/bin/env node
// Entry point for HTTP mode (`mstodo-http`). It reads .env from the working directory, then
// the environment, and refuses to start until every required setting is valid.
import dotenv from "dotenv"

import { readHttpConfig } from "./http-config.js"
import { startHttpServer } from "./http-server.js"

dotenv.config({ quiet: true })

readHttpConfig(process.env).fold(
  (problems) => {
    console.error("Cannot start the HTTP server:")
    for (const problem of problems) console.error(`  - ${problem}`)
    console.error("See the README's HTTP Mode section.")
    process.exit(1)
  },
  (config) => {
    startHttpServer(config).then(
      ({ server, provider }) => {
        // Stop taking requests and stop the provider's timers before exiting.
        const shutdown = () => {
          provider.close()
          server.stop().finally(() => process.exit(0))
        }
        process.once("SIGTERM", shutdown)
        process.once("SIGINT", shutdown)
      },
      (error) => {
        console.error("Error starting the HTTP server:", error instanceof Error ? error.message : String(error))
        process.exit(1)
      },
    )
  },
)
