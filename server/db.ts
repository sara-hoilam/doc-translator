import { eq, desc, gt, and } from "drizzle-orm";
import { drizzle } from "drizzle-orm/mysql2";
import { InsertUser, users, jobs, jobSteps, jobLogs, InsertJob, Job, JobStep, JobLog } from "../drizzle/schema";
import { ENV } from './_core/env';

let _db: ReturnType<typeof drizzle> | null = null;

export async function getDb() {
  if (!_db && process.env.DATABASE_URL) {
    try {
      _db = drizzle(process.env.DATABASE_URL);
    } catch (error) {
      console.warn("[Database] Failed to connect:", error);
      _db = null;
    }
  }
  return _db;
}

export async function upsertUser(user: InsertUser): Promise<void> {
  if (!user.openId) throw new Error("User openId is required for upsert");
  const db = await getDb();
  if (!db) { console.warn("[Database] Cannot upsert user: database not available"); return; }
  try {
    const values: InsertUser = { openId: user.openId };
    const updateSet: Record<string, unknown> = {};
    const textFields = ["name", "email", "loginMethod"] as const;
    type TextField = (typeof textFields)[number];
    const assignNullable = (field: TextField) => {
      const value = user[field];
      if (value === undefined) return;
      const normalized = value ?? null;
      values[field] = normalized;
      updateSet[field] = normalized;
    };
    textFields.forEach(assignNullable);
    if (user.lastSignedIn !== undefined) { values.lastSignedIn = user.lastSignedIn; updateSet.lastSignedIn = user.lastSignedIn; }
    if (user.role !== undefined) { values.role = user.role; updateSet.role = user.role; }
    else if (user.openId === ENV.ownerOpenId) { values.role = 'admin'; updateSet.role = 'admin'; }
    if (!values.lastSignedIn) values.lastSignedIn = new Date();
    if (Object.keys(updateSet).length === 0) updateSet.lastSignedIn = new Date();
    await db.insert(users).values(values).onDuplicateKeyUpdate({ set: updateSet });
  } catch (error) { console.error("[Database] Failed to upsert user:", error); throw error; }
}

// ── Anonymous user ────────────────────────────────────────────────────────────
// All jobs from unauthenticated sessions are owned by a single "anon" user row.
// We create it once and cache the id so userId = 0 is never written to the DB.

const ANON_OPEN_ID = "anon";
let _anonUserId: number | null = null;

export async function getAnonUserId(): Promise<number> {
  if (_anonUserId !== null) return _anonUserId;
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  // Try to find existing anon user
  const existing = await db.select({ id: users.id }).from(users)
    .where(eq(users.openId, ANON_OPEN_ID)).limit(1);
  if (existing.length > 0) {
    _anonUserId = existing[0].id;
    return _anonUserId;
  }

  // Create it if it doesn't exist yet
  const result = await db.insert(users).values({
    openId: ANON_OPEN_ID,
    name: "Anonymous",
    loginMethod: "none",
    lastSignedIn: new Date(),
  }).onDuplicateKeyUpdate({ set: { lastSignedIn: new Date() } });

  // Fetch the newly created (or re-found) id
  const created = await db.select({ id: users.id }).from(users)
    .where(eq(users.openId, ANON_OPEN_ID)).limit(1);
  _anonUserId = created[0].id;
  return _anonUserId;
}

export async function getUserByOpenId(openId: string) {
  const db = await getDb();
  if (!db) return undefined;
  const result = await db.select().from(users).where(eq(users.openId, openId)).limit(1);
  return result.length > 0 ? result[0] : undefined;
}

export async function getUserById(id: number) {
  const db = await getDb();
  if (!db) return undefined;
  const result = await db.select().from(users).where(eq(users.id, id)).limit(1);
  return result.length > 0 ? result[0] : undefined;
}

// ---- Jobs ----

export async function createJob(data: InsertJob): Promise<number> {
  const db = await getDb();
  if (!db) throw new Error("Database not available");
  const result = await db.insert(jobs).values(data);
  return (result[0] as any).insertId as number;
}

export async function getJobById(id: number): Promise<Job | undefined> {
  const db = await getDb();
  if (!db) return undefined;
  const result = await db.select().from(jobs).where(eq(jobs.id, id)).limit(1);
  return result[0];
}

export async function updateJob(id: number, data: Partial<Job>): Promise<void> {
  const db = await getDb();
  if (!db) throw new Error("Database not available");
  await db.update(jobs).set(data as any).where(eq(jobs.id, id));
}

export async function getUserJobs(userId: number): Promise<Job[]> {
  const db = await getDb();
  if (!db) return [];
  return db.select().from(jobs).where(eq(jobs.userId, userId)).orderBy(desc(jobs.createdAt)).limit(50);
}

/** Check if a job has been cancelled (fast DB read) */
export async function isJobCancelled(id: number): Promise<boolean> {
  const db = await getDb();
  if (!db) return false;
  const result = await db.select({ cancelled: jobs.cancelled }).from(jobs).where(eq(jobs.id, id)).limit(1);
  return result[0]?.cancelled === true || (result[0]?.cancelled as any) === 1;
}

/** Mark job as cancelled */
export async function cancelJob(id: number): Promise<void> {
  const db = await getDb();
  if (!db) throw new Error("Database not available");
  await db.update(jobs).set({ cancelled: true, status: "cancelled" } as any).where(eq(jobs.id, id));
}

// ---- Job Steps ----

export async function createJobStep(jobId: number, step: JobStep["step"]): Promise<void> {
  const db = await getDb();
  if (!db) return;
  await db.insert(jobSteps).values({ jobId, step, status: "pending" });
}

export async function updateJobStep(jobId: number, step: JobStep["step"], data: Partial<JobStep>): Promise<void> {
  const db = await getDb();
  if (!db) return;
  await db.update(jobSteps).set(data as any).where(eq(jobSteps.jobId, jobId));
}

export async function getJobSteps(jobId: number): Promise<JobStep[]> {
  const db = await getDb();
  if (!db) return [];
  return db.select().from(jobSteps).where(eq(jobSteps.jobId, jobId));
}

// ---- Job Logs ----

// In-memory sequence counter per job (avoids DB round-trip for seq)
const seqCounters = new Map<number, number>();

export async function appendJobLog(
  jobId: number,
  level: JobLog["level"],
  message: string,
): Promise<void> {
  const db = await getDb();
  if (!db) return;
  const seq = (seqCounters.get(jobId) ?? 0) + 1;
  seqCounters.set(jobId, seq);
  await db.insert(jobLogs).values({ jobId, seq, level, message, createdAt: Date.now() });
}

export async function getJobLogs(jobId: number, afterSeq = 0): Promise<JobLog[]> {
  const db = await getDb();
  if (!db) return [];
  return db
    .select()
    .from(jobLogs)
    .where(and(eq(jobLogs.jobId, jobId), gt(jobLogs.seq, afterSeq)))
    .orderBy(jobLogs.seq)
    .limit(200);
}

export async function clearJobLogs(jobId: number): Promise<void> {
  const db = await getDb();
  if (!db) return;
  await db.delete(jobLogs).where(eq(jobLogs.jobId, jobId));
  seqCounters.delete(jobId);
}

/** Permanently delete a job and all its steps/logs from the DB */
export async function deleteJob(id: number): Promise<void> {
  const db = await getDb();
  if (!db) return;
  await db.delete(jobLogs).where(eq(jobLogs.jobId, id));
  await db.delete(jobSteps).where(eq(jobSteps.jobId, id));
  await db.delete(jobs).where(eq(jobs.id, id));
  seqCounters.delete(id);
}
