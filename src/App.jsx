import React, { useEffect, useState } from "react";
import { supabase } from "./supabaseClient";
import Login from "./components/Login.jsx";
import DraftBoard from "./components/DraftBoard.jsx";

export default function App() {
    const [session, setSession] = useState(undefined); // undefined = loading, null = signed out

  useEffect(() => {
        supabase.auth.getSession().then(({ data }) => setSession(data.session));
        const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => {
                setSession(newSession);
        });
        return () => listener.subscription.unsubscribe();
  }, []);

  if (session === undefined) {
        return (
                <div style={{ background: "#15171A", color: "#9C9C93", minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'IBM Plex Mono', monospace", fontSize: "13px" }}>
                          Loading…
                </div>div>
              );
  }

  if (!session) return <Login />;

  return <DraftBoard session={session} />;
}
