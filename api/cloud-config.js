module.exports = function cloudConfigHandler(_req, res) {
  const config = {
    url: process.env.SUPABASE_URL || "",
    anonKey: process.env.SUPABASE_ANON_KEY || "",
    table: process.env.SUPABASE_TABLE || "dashboard_state",
    stateKey: process.env.SUPABASE_STATE_KEY || "global",
    autoSync: String(process.env.SUPABASE_AUTO_SYNC || "true").toLowerCase() !== "false",
  };

  res.setHeader("Content-Type", "application/javascript; charset=utf-8");
  res.setHeader("Cache-Control", "s-maxage=60, stale-while-revalidate=600");
  res.statusCode = 200;
  res.end(`window.__HRBD_CLOUD_CONFIG__ = ${JSON.stringify(config)};`);
};
