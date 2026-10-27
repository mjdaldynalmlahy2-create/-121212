import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import admin from 'firebase-admin';
import { getFirestore } from 'firebase-admin/firestore';
import firebaseConfig from './firebase-applet-config.json' with { type: 'json' };

const firebaseAdminApp = admin.initializeApp({
  projectId: firebaseConfig.projectId,
});

async function startServer() {
  const app = express();
  const PORT = 3000;

  // JSON middleware
  app.use(express.json());

  // API routines
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  app.post('/api/admin/reset-password', async (req, res) => {
    const { uid, newPassword, idToken } = req.body;
    console.log(`[Admin API] Password reset request for: ${uid}`);
    
    try {
      if (!idToken) throw new Error('Missing ID token');

      // 1. Verify the requester's identity
      const decodedToken = await admin.auth().verifyIdToken(idToken);
      const requesterUid = decodedToken.uid;

      // 2. Check if the requester is an admin
      const db = getFirestore(firebaseAdminApp, firebaseConfig.firestoreDatabaseId);
      const userDoc = await db.collection('users').doc(requesterUid).get();
      const userData = userDoc.data();

      if (!userData || userData.role !== 'admin') {
        console.warn(`[Admin API] Unauthorized attempt from: ${requesterUid}`);
        return res.status(403).json({ error: 'غير مسموح. يجب أن تكون مديراً للنظام.' });
      }

      // 3. Reset the target user's password
      console.log(`[Admin API] Updating password for ${uid}...`);
      await admin.auth().updateUser(uid, { password: String(newPassword || '123456') });
      
      console.log(`[Admin API] Success for ${uid}`);
      res.json({ success: true, message: 'تم إعادة تعيين كلمة المرور بنجاح' });
    } catch (error: any) {
      console.error('[Admin API] Error:', error);
      res.status(500).json({ error: error.message || 'حدث خطأ في الخادم' });
    }
  });

  app.post('/api/admin/delete-user', async (req, res) => {
    const { uid, idToken } = req.body;
    console.log(`[Admin API] User deletion request for: ${uid}`);
    
    try {
      if (!idToken) throw new Error('Missing ID token');

      // 1. Verify the requester's identity
      const decodedToken = await admin.auth().verifyIdToken(idToken);
      const requesterUid = decodedToken.uid;

      // 2. Check if the requester is an admin
      const db = getFirestore(firebaseAdminApp, firebaseConfig.firestoreDatabaseId);
      const userDoc = await db.collection('users').doc(requesterUid).get();
      const userData = userDoc.data();

      if (!userData || userData.role !== 'admin') {
        console.warn(`[Admin API] Unauthorized deletion attempt from: ${requesterUid}`);
        return res.status(403).json({ error: 'غير مسموح. يجب أن تكون مديراً للنظام.' });
      }

      // 3. Delete the user from Auth and Firestore
      console.log(`[Admin API] Deleting user ${uid}...`);
      
      let authDeleted = false;
      let firestoreDeleted = false;

      try {
        await admin.auth().deleteUser(uid);
        authDeleted = true;
      } catch (authErr: any) {
        console.warn(`[Admin API] Auth deletion failed for ${uid}:`, authErr.message);
      }

      try {
        await db.collection('users').doc(uid).delete();
        firestoreDeleted = true;
      } catch (fsErr: any) {
        console.warn(`[Admin API] Firestore deletion failed for ${uid}:`, fsErr.message);
      }
      
      if (!authDeleted && !firestoreDeleted) {
        throw new Error('فشل حذف المستخدم من جميع المصادر');
      }

      console.log(`[Admin API] Successfully handled deletion for user ${uid}. Auth: ${authDeleted}, FS: ${firestoreDeleted}`);
      res.json({ success: true, message: 'تم حذف المستخدم بنجاح' });
    } catch (error: any) {
      console.error('[Admin API] Deletion Error:', error);
      res.status(500).json({ error: error.message || 'حدث خطأ في الخادم' });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
