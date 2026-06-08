import { join } from "@std/path";

// Read tokens from file
const TOKEN_FILE_PATH = Deno.env.get("MSTODO_TOKEN_FILE") || join(Deno.cwd(), "tokens.json");
const MS_GRAPH_BASE = "https://graph.microsoft.com/v1.0";

console.log("🔍 Microsoft Graph API Explorer for To Do\n" + "=".repeat(50) + "\n");

let accessToken: string;
try {
  const tokenData = JSON.parse(Deno.readTextFileSync(TOKEN_FILE_PATH));
  accessToken = tokenData.accessToken;
  console.log("✅ Token loaded successfully\n");
} catch {
  console.error("❌ Failed to load tokens from:", TOKEN_FILE_PATH);
  console.error('Please run "deno task auth" first');
  Deno.exit(1);
}

async function makeRequest(url: string, token: string): Promise<unknown> {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`HTTP ${response.status}: ${error}`);
  }

  return response.json();
}

async function exploreAPI() {
  console.log("📊 Test 1: Using $select=* to get all properties\n");
  try {
    const lists = (await makeRequest(`${MS_GRAPH_BASE}/me/todo/lists?$select=*`, accessToken)) as {
      value: Record<string, unknown>[];
    };
    if (lists.value && lists.value.length > 0) {
      const properties = Object.keys(lists.value[0]);
      console.log(`Found ${properties.length} properties:`, properties.join(", "));
      console.log("\nFirst list example:");
      console.log(JSON.stringify(lists.value[0], null, 2));
    }
  } catch (error: unknown) {
    console.error("Error:", (error as Error).message);
  }

  console.log("\n" + "-".repeat(50) + "\n");
  console.log("📊 Test 2: Testing $expand options\n");

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
  ];

  for (const expand of expandOptions) {
    try {
      const response = (await makeRequest(
        `${MS_GRAPH_BASE}/me/todo/lists?$expand=${expand}&$top=1`,
        accessToken,
      )) as {
        value: Record<string, unknown>[];
      };

      if (response.value && response.value.length > 0 && response.value[0][expand]) {
        console.log(`✓ $expand=${expand}: Found data!`);
        console.log(JSON.stringify(response.value[0][expand], null, 2));
      } else {
        console.log(`✗ $expand=${expand}: No data returned`);
      }
    } catch (error: unknown) {
      console.log(`✗ $expand=${expand}: ${(error as Error).message.split("\n")[0]}`);
    }
  }

  console.log("\n" + "-".repeat(50) + "\n");
  console.log("📊 Test 3: Checking for folder/group endpoints\n");

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
  ];

  for (const endpoint of endpoints) {
    try {
      const response = await makeRequest(`${MS_GRAPH_BASE}${endpoint}`, accessToken);
      console.log(
        `✓ ${endpoint}: FOUND! Response:`,
        JSON.stringify(response).substring(0, 200) + "...",
      );
    } catch (error: unknown) {
      const msg = (error as Error).message.split(":")[1]?.trim() || (error as Error).message;
      console.log(`✗ ${endpoint}: ${msg}`);
    }
  }

  console.log("\n" + "-".repeat(50) + "\n");
  console.log("📊 Test 4: Checking list extensions\n");

  try {
    const lists = (await makeRequest(`${MS_GRAPH_BASE}/me/todo/lists?$top=1`, accessToken)) as {
      value: { id: string; displayName: string }[];
    };
    if (lists.value && lists.value.length > 0) {
      const listId = lists.value[0].id;
      console.log(`Testing extensions for list: ${lists.value[0].displayName}`);

      const extEndpoints = [
        `/me/todo/lists/${listId}/extensions`,
        `/me/todo/lists/${listId}/openExtensions`,
        `/me/todo/lists/${listId}?$expand=extensions`,
        `/me/todo/lists/${listId}?$expand=singleValueExtendedProperties`,
      ];

      for (const endpoint of extEndpoints) {
        try {
          const response = await makeRequest(`${MS_GRAPH_BASE}${endpoint}`, accessToken);
          console.log(`✓ ${endpoint}: Found data:`, JSON.stringify(response, null, 2));
        } catch (error: unknown) {
          const msg = (error as Error).message.split(":")[1]?.trim() || (error as Error).message;
          console.log(`✗ ${endpoint}: ${msg}`);
        }
      }
    }
  } catch (error: unknown) {
    console.error("Error getting lists:", (error as Error).message);
  }

  console.log("\n" + "=".repeat(50) + "\n");
  console.log("✅ API exploration complete!");

  // Additional beta endpoint test
  console.log("\n📊 Bonus: Testing beta endpoints\n");
  const BETA_BASE = "https://graph.microsoft.com/beta";

  try {
    const betaLists = (await makeRequest(`${BETA_BASE}/me/todo/lists`, accessToken)) as {
      value: Record<string, unknown>[];
    };
    if (betaLists.value && betaLists.value.length > 0) {
      const betaProperties = Object.keys(betaLists.value[0]);
      console.log(`Beta API properties (${betaProperties.length}):`, betaProperties.join(", "));

      const v1Lists = (await makeRequest(`${MS_GRAPH_BASE}/me/todo/lists?$top=1`, accessToken)) as {
        value: Record<string, unknown>[];
      };
      const v1Properties = Object.keys(v1Lists.value[0]);

      const newProperties = betaProperties.filter((p) => !v1Properties.includes(p));
      if (newProperties.length > 0) {
        console.log("\n🎉 New properties in beta:", newProperties.join(", "));
        console.log("\nBeta list example:");
        console.log(JSON.stringify(betaLists.value[0], null, 2));
      } else {
        console.log("\nNo additional properties found in beta API");
      }
    }
  } catch (error: unknown) {
    console.error("Beta API error:", (error as Error).message);
  }
}

// Run the exploration
exploreAPI().catch((error: unknown) => {
  console.error("\n❌ Fatal error:", error);
  Deno.exit(1);
});
