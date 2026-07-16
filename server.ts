import express from 'express';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const DB_DIR = path.join(process.cwd(), 'data');
const DB_FILE = path.join(DB_DIR, 'dashboards.json');

// Initialize database
function initDb() {
  if (!fs.existsSync(DB_DIR)) {
    fs.mkdirSync(DB_DIR, { recursive: true });
  }
  if (!fs.existsSync(DB_FILE)) {
    fs.writeFileSync(DB_FILE, JSON.stringify({ dashboards: {} }, null, 2), 'utf-8');
  }
}

function readDb() {
  try {
    initDb();
    const content = fs.readFileSync(DB_FILE, 'utf-8');
    return JSON.parse(content);
  } catch (err) {
    console.error('Error reading database, resetting:', err);
    return { dashboards: {} };
  }
}

function writeDb(data: any) {
  try {
    initDb();
    fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (err) {
    console.error('Error writing database:', err);
  }
}

// Generate unique ID
function generateId() {
  return 'sh_' + Math.random().toString(36).substring(2, 8);
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Middleware for parsing JSON with a larger limit (to accommodate base64 logos and report datasets)
  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ limit: '50mb', extended: true }));

  // API Endpoints

  // Publish a report under a shareId and date
  app.post('/api/dashboard/publish', (req, res) => {
    try {
      const { shareId, date, logo, reportData, grandTotals, selectedTeams } = req.body;

      if (!date || !reportData || !grandTotals) {
        return res.status(400).json({ error: 'Missing required report fields (date, reportData, grandTotals).' });
      }

      const db = readDb();
      let targetShareId = shareId;

      if (!targetShareId || !db.dashboards[targetShareId]) {
        targetShareId = generateId();
        db.dashboards[targetShareId] = {
          shareId: targetShareId,
          logo: logo || null,
          dates: {},
          updatedAt: new Date().toISOString()
        };
      }

      // If a new logo is uploaded during publish, update it in the workspace
      if (logo) {
        db.dashboards[targetShareId].logo = logo;
      }

      // Store the specific date report
      db.dashboards[targetShareId].dates[date] = {
        reportData,
        grandTotals,
        selectedTeams: selectedTeams || [],
        updatedAt: new Date().toISOString()
      };

      db.dashboards[targetShareId].updatedAt = new Date().toISOString();
      writeDb(db);

      res.json({
        success: true,
        shareId: targetShareId,
        date
      });
    } catch (error: any) {
      console.error('Publish error:', error);
      res.status(500).json({ error: 'Failed to publish dashboard: ' + error.message });
    }
  });

  // Get metadata (logo & available dates) for a shareId
  app.get('/api/dashboard/share/:shareId', (req, res) => {
    try {
      const { shareId } = req.params;
      const db = readDb();
      const workspace = db.dashboards[shareId];

      if (!workspace) {
        return res.status(404).json({ error: 'Dashboard not found.' });
      }

      const availableDates = Object.keys(workspace.dates).sort((a, b) => b.localeCompare(a)); // Newest first

      res.json({
        success: true,
        shareId,
        logo: workspace.logo,
        availableDates
      });
    } catch (error: any) {
      res.status(500).json({ error: 'Failed to load sharing info: ' + error.message });
    }
  });

  // Get specific report data for a shareId and date
  app.get('/api/dashboard/share/:shareId/date/:date', (req, res) => {
    try {
      const { shareId, date } = req.params;
      const db = readDb();
      const workspace = db.dashboards[shareId];

      if (!workspace) {
        return res.status(404).json({ error: 'Dashboard not found.' });
      }

      const report = workspace.dates[date];
      if (!report) {
        return res.status(404).json({ error: 'Report data not found for the specified date.' });
      }

      res.json({
        success: true,
        shareId,
        date,
        logo: workspace.logo, // Fallback to workspace logo in case it's needed
        reportData: report.reportData,
        grandTotals: report.grandTotals,
        selectedTeams: report.selectedTeams
      });
    } catch (error: any) {
      res.status(500).json({ error: 'Failed to load report data: ' + error.message });
    }
  });

  // Vite Integration & Asset Serving
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
