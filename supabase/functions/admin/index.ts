// Owner-only admin actions: list all users, view a user's sadhana data, reset a user's password.
// Deploy with: supabase functions deploy admin
// Requires these function secrets (set via `supabase secrets set`):
//   SUPABASE_URL              - your project URL (same as VITE_SUPABASE_URL)
//   SUPABASE_SERVICE_ROLE_KEY - service role key from Project Settings > API (NEVER put this in the client app)
//   OWNER_EMAIL                - the email address allowed to call this function

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });

  const json = (body, status = 200) =>
    new Response(JSON.stringify(body), { status, headers: { ...CORS_HEADERS, "Content-Type": "application/json" } });

  try {
    const authHeader = req.headers.get("Authorization") || "";
    const jwt = authHeader.replace(/^Bearer\s+/i, "");
    if (!jwt) return json({ error: "Missing Authorization header" }, 401);

    const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
    const SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    const OWNER_EMAIL = (Deno.env.get("OWNER_EMAIL") || "").toLowerCase();

    // Client scoped to the caller's own token, just to find out who they are.
    const callerClient = createClient(SUPABASE_URL, SERVICE_ROLE_KEY, {
      global: { headers: { Authorization: `Bearer ${jwt}` } },
    });
    const { data: callerData, error: callerErr } = await callerClient.auth.getUser(jwt);
    if (callerErr || !callerData?.user) return json({ error: "Invalid session" }, 401);
    if ((callerData.user.email || "").toLowerCase() !== OWNER_EMAIL) {
      return json({ error: "Not authorized" }, 403);
    }

    // Privileged client - only reached once we've confirmed the caller is the owner.
    const admin = createClient(SUPABASE_URL, SERVICE_ROLE_KEY);
    const { action, userId, newPassword } = await req.json();

    if (action === "listUsers") {
      const { data, error } = await admin.auth.admin.listUsers();
      if (error) throw error;
      const users = data.users.map((u) => ({
        id: u.id, email: u.email, createdAt: u.created_at, lastSignInAt: u.last_sign_in_at,
      }));
      return json({ users });
    }

    if (action === "getUserData") {
      if (!userId) return json({ error: "userId required" }, 400);
      const { data, error } = await admin.from("sadhana_app_state").select("data").eq("user_id", userId).maybeSingle();
      if (error) throw error;
      return json({ data: data?.data || null });
    }

    if (action === "resetPassword") {
      if (!userId || !newPassword) return json({ error: "userId and newPassword required" }, 400);
      if (newPassword.length < 6) return json({ error: "Password must be at least 6 characters" }, 400);
      const { error } = await admin.auth.admin.updateUserById(userId, { password: newPassword });
      if (error) throw error;
      return json({ ok: true });
    }

    return json({ error: "Unknown action" }, 400);
  } catch (e) {
    return json({ error: e.message || "Server error" }, 500);
  }
});
