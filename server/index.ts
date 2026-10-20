/**
 * CA Tools Backend Automation Server Entry Point
 */

import { app } from './app';

const PORT = process.env.PORT || 5000;

app.listen(PORT, () => {
  console.log(`================================================================`);
  console.log(`🚀 CA Tools Intelligent Backend Automation Engine LIVE on :${PORT}`);
  console.log(`📡 API Endpoints available at http://localhost:${PORT}/api`);
  console.log(`🛡️  ACID Data Store & Firm-Level Isolation Active`);
  console.log(`================================================================`);
});
