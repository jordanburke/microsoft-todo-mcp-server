#!/usr/bin/env node
import { readFileSync } from "fs"
import { join } from "path"

// Read tokens from file
const TOKEN_FILE_PATH = process.env.MSTODO_TOKEN_FILE || join(process.cwd(), "tokens.json")
const MS_GRAPH_BASE = "https://graph.microsoft.com/v1.0"

console.log("🔍 Microsoft Graph API Explorer for To Do\n" + "=".repeat(50) + "\n")

let accessToken
try {
  const tokenData = JSON.parse(readFileSync(TOKEN_FILE_PATH, "utf8"))
  accessToken = tokenData.accessToken
  console.log("✅ Token loaded successfully\n")
} catch (error) {
  console.error("❌ Failed to load tokens from:", TOKEN_FILE_PATH)
  console.error('Please run "pnpm run auth" first')
  process.exit(1)
}

async function makeRequest(url, token) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  })

  if (!response.ok) {
    const error = await response.text()
    throw new Error(`HTTP ${response.status}: ${error}`)
  }

  return response.json()
}

async function exploreAPI() {
  console.log("📊 Test 1: Using $select=* to get all properties\n")
  try {
    const lists = await makeRequest(`${MS_GRAPH_BASE}/me/todo/lists?$select=*`, accessToken)
    if (lists.value && lists.value.length > 0) {
      const properties = Object.keys(lists.value[0])
      console.log(`Found ${properties.length} properties:`, properties.join(", "))
      console.log("\nFirst list example:")
      console.log(JSON.stringify(lists.value[0], null, 2))
    }
  } catch (error) {
    console.error("Error:", error.message)
  }

  console.log("\n" + "-".repeat(50) + "\n")
  console.log("📊 Test 2: Testing $expand options\n")

  const expandOptions = [
    "extensions",
    "singleValueExtendedProperties",
    "multiValueExtendedProperties",
    "openExtensions",
    "parent",
    "children",
    "folder",
    "parentFolder",
    "group",
    "category",
    "parentReference",
    "childFolders",
  ]

  for (const expand of expandOptions) {
    try {
      const response = await makeRequest(`${MS_GRAPH_BASE}/me/todo/lists?$expand=${expand}&$top=1`, accessToken)

      if (response.value && response.value.length > 0 && response.value[0][expand]) {
        console.log(`✓ $expand=${expand}: Found data!`)
        console.log(JSON.stringify(response.value[0][expand], null, 2))
      } else {
        console.log(`✗ $expand=${expand}: No data returned`)
      }
    } catch (error) {
      console.log(`✗ $expand=${expand}: ${error.message.split("\n")[0]}`)
    }
  }

  console.log("\n" + "-".repeat(50) + "\n")
  console.log("📊 Test 3: Checking for folder/group endpoints\n")

  const endpoints = [
    "/me/todo/folders",
    "/me/todo/groups",
    "/me/todo/listGroups",
    "/me/todo/listFolders",
    "/me/todo/categories",
    "/me/todo/lists/folders",
    "/me/todo/lists/groups",
    "/me/outlook/taskGroups",
    "/me/outlook/taskFolders",
  ]

  for (const endpoint of endpoints) {
    try {
      const response = await makeRequest(`${MS_GRAPH_BASE}${endpoint}`, accessToken)
      console.log(`✓ ${endpoint}: FOUND! Response:`, JSON.stringify(response).substring(0, 200) + "...")
    } catch (error) {
      console.log(`✗ ${endpoint}: ${error.message.split(":")[1].trim()}`)
    }
  }

  console.log("\n" + "-".repeat(50) + "\n")
  console.log("📊 Test 4: Checking list extensions\n")

  try {
    const lists = await makeRequest(`${MS_GRAPH_BASE}/me/todo/lists?$top=1`, accessToken)
    if (lists.value && lists.value.length > 0) {
      const listId = lists.value[0].id
      console.log(`Testing extensions for list: ${lists.value[0].displayName}`)

      // Try different extension endpoints
      const extEndpoints = [
        `/me/todo/lists/${listId}/extensions`,
        `/me/todo/lists/${listId}/openExtensions`,
        `/me/todo/lists/${listId}?$expand=extensions`,
        `/me/todo/lists/${listId}?$expand=singleValueExtendedProperties`,
      ]

      for (const endpoint of extEndpoints) {
        try {
          const response = await makeRequest(`${MS_GRAPH_BASE}${endpoint}`, accessToken)
          console.log(`✓ ${endpoint}: Found data:`, JSON.stringify(response, null, 2))
        } catch (error) {
          console.log(`✗ ${endpoint}: ${error.message.split(":")[1].trim()}`)
        }
      }
    }
  } catch (error) {
    console.error("Error getting lists:", error.message)
  }

  console.log("\n" + "=".repeat(50) + "\n")
  console.log("✅ API exploration complete!")

  // Additional beta endpoint test
  console.log("\n📊 Bonus: Testing beta endpoints\n")
  const BETA_BASE = "https://graph.microsoft.com/beta"

  try {
    const betaLists = await makeRequest(`${BETA_BASE}/me/todo/lists`, accessToken)
    if (betaLists.value && betaLists.value.length > 0) {
      const betaProperties = Object.keys(betaLists.value[0])
      console.log(`Beta API properties (${betaProperties.length}):`, betaProperties.join(", "))

      // Check if beta has different properties
      const v1Lists = await makeRequest(`${MS_GRAPH_BASE}/me/todo/lists?$top=1`, accessToken)
      const v1Properties = Object.keys(v1Lists.value[0])

      const newProperties = betaProperties.filter((p) => !v1Properties.includes(p))
      if (newProperties.length > 0) {
        console.log("\n🎉 New properties in beta:", newProperties.join(", "))
        console.log("\nBeta list example:")
        console.log(JSON.stringify(betaLists.value[0], null, 2))
      } else {
        console.log("\nNo additional properties found in beta API")
      }
    }
  } catch (error) {
    console.error("Beta API error:", error.message)
  }
}

// Run the exploration
exploreAPI().catch((error) => {
  console.error("\n❌ Fatal error:", error)
  process.exit(1)
})
