import { describe, expect, it } from "vitest";
import { appRouter } from "./routers";
import type { TrpcContext } from "./_core/context";

function createPublicCtx(): TrpcContext {
  return {
    user: null,
    req: { protocol: "https", headers: {} } as TrpcContext["req"],
    res: { clearCookie: () => {} } as TrpcContext["res"],
  };
}

function createAuthCtx(userId = 1): TrpcContext {
  return {
    user: {
      id: userId,
      openId: "test-user",
      email: "test@example.com",
      name: "Test User",
      loginMethod: "manus",
      role: "user",
      createdAt: new Date(),
      updatedAt: new Date(),
      lastSignedIn: new Date(),
    },
    req: { protocol: "https", headers: {} } as TrpcContext["req"],
    res: { clearCookie: () => {} } as TrpcContext["res"],
  };
}

describe("docs.config", () => {
  it("returns supported formats and languages", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    const config = await caller.docs.config();
    expect(config.inputFormats).not.toContain("pdf"); // PDF removed as input format
    expect(config.inputFormats).toContain("docx");
    expect(config.outputFormats).toContain("pdf");
    expect(config.outputFormats).toContain("csv");
    expect(config.languages.length).toBeGreaterThan(50);
    expect(config.costPerPageExtraction).toBeGreaterThan(0);
    expect(config.costPerPageTranslation).toBeGreaterThan(0);
  });
});

describe("docs.estimate", () => {
  it("calculates extraction-only cost", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    const est = await caller.docs.estimate({ pageCount: 5, doTranslate: false, doConvert: false });
    expect(est.extractionCost).toBeGreaterThan(0);
    expect(est.translationCost).toBe(0);
    expect(est.total).toBeCloseTo(est.extractionCost, 5);
  });

  it("calculates extraction + translation cost", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    const est = await caller.docs.estimate({ pageCount: 10, doTranslate: true, doConvert: false });
    expect(est.extractionCost).toBeGreaterThan(0);
    expect(est.translationCost).toBeGreaterThan(0);
    expect(est.total).toBeCloseTo(est.extractionCost + est.translationCost, 5);
    expect(est.perPage).toBeCloseTo(est.total / 10, 3);
  });

  it("scales linearly with page count", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    const est1 = await caller.docs.estimate({ pageCount: 1, doTranslate: true, doConvert: false });
    const est10 = await caller.docs.estimate({ pageCount: 10, doTranslate: true, doConvert: false });
    expect(est10.total).toBeCloseTo(est1.total * 10, 3);
  });
});

describe("docs.cleanup", () => {
  it("returns success:false for non-existent job (public, no auth required)", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    const result = await caller.docs.cleanup({ jobId: 999999 });
    expect(result.success).toBe(false);
  });
});

describe("docs.status", () => {
  it("throws for non-existent job (public, no auth required)", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    await expect(caller.docs.status({ jobId: 999999 })).rejects.toThrow("Job not found");
  });
});

describe("docs.cancel", () => {
  it("throws for non-existent job (public, no auth required)", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    await expect(caller.docs.cancel({ jobId: 999999 })).rejects.toThrow("Job not found");
  });
});

describe("docs.getLogs", () => {
  it("throws for non-existent job (public, no auth required)", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    await expect(caller.docs.getLogs({ jobId: 999999, afterSeq: 0 })).rejects.toThrow("Job not found");
  });
});

describe("docs.resume", () => {
  it("throws for non-existent job (public, no auth required)", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    await expect(caller.docs.resume({ jobId: 999999 })).rejects.toThrow("Job not found");
  });
});

describe("docs.killPaused", () => {
  it("throws for non-existent job (public, no auth required)", async () => {
    const caller = appRouter.createCaller(createPublicCtx());
    await expect(caller.docs.killPaused({ jobId: 999999 })).rejects.toThrow("Job not found");
  });
});

describe("auth.logout", () => {
  it("clears session cookie", async () => {
    const cleared: string[] = [];
    const ctx: TrpcContext = {
      user: {
        id: 1, openId: "u", email: null, name: null, loginMethod: null,
        role: "user", createdAt: new Date(), updatedAt: new Date(), lastSignedIn: new Date(),
      },
      req: { protocol: "https", headers: {} } as TrpcContext["req"],
      res: { clearCookie: (name: string) => cleared.push(name) } as TrpcContext["res"],
    };
    const caller = appRouter.createCaller(ctx);
    const result = await caller.auth.logout();
    expect(result.success).toBe(true);
    expect(cleared).toHaveLength(1);
  });
});
