/**
 * oauth.ts — OAuth callback route
 *
 * The app currently operates without requiring sign-in (all jobs run as
 * anonymous userId = 0). This file registers a no-op placeholder so the
 * rest of the server code that calls registerOAuthRoutes() continues to
 * compile and run.
 *
 * If you want to add Google Sign-In in the future, implement it here using
 * Google's OAuth 2.0 flow and the existing JWT session helpers in sdk.ts.
 */

import type { Express } from "express";

// eslint-disable-next-line @typescript-eslint/no-unused-vars
export function registerOAuthRoutes(_app: Express): void {
  // No-op: sign-in is not required in the current version.
}
