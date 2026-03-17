import { int, mysqlEnum, mysqlTable, text, timestamp, varchar, float, boolean, bigint } from "drizzle-orm/mysql-core";

export const users = mysqlTable("users", {
  id: int("id").autoincrement().primaryKey(),
  openId: varchar("openId", { length: 64 }).notNull().unique(),
  name: text("name"),
  email: varchar("email", { length: 320 }),
  loginMethod: varchar("loginMethod", { length: 64 }),
  role: mysqlEnum("role", ["user", "admin"]).default("user").notNull(),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
  lastSignedIn: timestamp("lastSignedIn").defaultNow().notNull(),
});

export type User = typeof users.$inferSelect;
export type InsertUser = typeof users.$inferInsert;

export const jobs = mysqlTable("jobs", {
  id: int("id").autoincrement().primaryKey(),
  userId: int("userId").notNull(),
  originalFileName: varchar("originalFileName", { length: 512 }).notNull(),
  originalFileKey: varchar("originalFileKey", { length: 512 }).notNull(),
  originalFileUrl: text("originalFileUrl").notNull(),
  originalFormat: varchar("originalFormat", { length: 32 }).notNull(),
  outputFormat: varchar("outputFormat", { length: 32 }),
  sourceLanguage: varchar("sourceLanguage", { length: 16 }),
  targetLanguage: varchar("targetLanguage", { length: 16 }),
  targetLanguageName: varchar("targetLanguageName", { length: 64 }),
  pageCount: int("pageCount").default(0),
  charCount: int("charCount").default(0),
  estimatedCost: float("estimatedCost").default(0),
  conversionCostUsd: float("conversionCostUsd").default(0),
  downloadPriceUsd: float("downloadPriceUsd").default(0),
  paid: boolean("paid").default(false).notNull(),
  stripeSessionId: varchar("stripeSessionId", { length: 256 }),
  status: mysqlEnum("status", ["pending", "extracting", "translating", "converting", "done", "error", "cancelled", "paused"]).default("pending").notNull(),
  cancelled: boolean("cancelled").default(false).notNull(),
  errorMessage: text("errorMessage"),
  outputFileKey: varchar("outputFileKey", { length: 512 }),
  outputFileUrl: text("outputFileUrl"),
  previewFileKey: varchar("previewFileKey", { length: 512 }),
  previewFileUrl: text("previewFileUrl"),
  previewPageCount: int("previewPageCount").default(0),
  extractedText: text("extractedText"),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
});

export type Job = typeof jobs.$inferSelect;
export type InsertJob = typeof jobs.$inferInsert;

export const jobSteps = mysqlTable("job_steps", {
  id: int("id").autoincrement().primaryKey(),
  jobId: int("jobId").notNull(),
  step: mysqlEnum("step", ["upload", "extract", "translate", "convert"]).notNull(),
  status: mysqlEnum("status", ["pending", "running", "done", "error"]).default("pending").notNull(),
  message: text("message"),
  startedAt: timestamp("startedAt"),
  completedAt: timestamp("completedAt"),
});

export type JobStep = typeof jobSteps.$inferSelect;

export const jobLogs = mysqlTable("job_logs", {
  id: int("id").autoincrement().primaryKey(),
  jobId: int("jobId").notNull(),
  seq: int("seq").notNull(),
  level: mysqlEnum("level", ["info", "progress", "success", "warning", "error"]).default("info").notNull(),
  message: text("message").notNull(),
  createdAt: bigint("createdAt", { mode: "number" }).notNull(), // unix ms for fast ordering
});

export type JobLog = typeof jobLogs.$inferSelect;
