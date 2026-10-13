import { defineConfig } from "drizzle-kit";

// `drizzle-kit generate` works offline; only migrate/push need DATABASE_URL.
const connectionString = process.env.DATABASE_URL;

export default defineConfig({
  schema: "./drizzle/schema.ts",
  out: "./drizzle",
  dialect: "mysql",
  ...(connectionString ? { dbCredentials: { url: connectionString } } : {}),
});
