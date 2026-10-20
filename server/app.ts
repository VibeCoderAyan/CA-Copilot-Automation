/**
 * Express Application Setup for CA Tools Backend Automation Engine
 */

import express from 'express';
import type { Request, Response, NextFunction } from 'express';
import cors from 'cors';
import { apiRouter } from './routes/apiRoutes.ts';

export const app = express();

// Middlewares
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'x-firm-id', 'x-firm-name', 'x-user-id', 'x-user-role', 'x-client-id'],
}));

// Support large file payloads (up to 50MB for batch invoices and ledgers)
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Mount API Router
app.use('/api', apiRouter);

// Global Error Handler
app.use((err: any, _req: Request, res: Response, _next: NextFunction) => {
  console.error('[API Error]:', err);
  res.status(err.status || 500).json({
    success: false,
    error: err.message || 'Internal Server Error',
    timestamp: new Date().toISOString(),
  });
});
