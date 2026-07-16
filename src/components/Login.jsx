import React, { useState } from "react";
import { supabase } from "../supabaseClient";

const COLORS = {
  bg: "#15171A",
  surface: "#1D2024",
  ink: "#ECE7DC",
  inkDim: "#9C9C93",
  hair: "rgba(236,231,220,0.10)",
  accent: "#C98A3E",
};

export default function Login() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [resetOpen, setResetOpen] = useState(false);
  const [resetEmail, setResetEmail] = useState("");
  const [resetMsg, setResetMsg] = useState("");
  const [resetLoading, setResetLoading] = useState(false);

  async function handleSignIn(e) {
    e.preventDefault();
    setError("");
    setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (error) setError(error.message);
  }

  async function handleResetPassword(e) {
    e.preventDefault();
    setResetMsg("");
    if (!resetEmail.trim()) return;
    setResetLoading(true);
    const { error } = await supabase.auth.resetPasswordForEmail(resetEmail.trim(), {
      redirectTo: window.location.origin,
    });
    setResetLoading(false);
    if (error) {
      setResetMsg(error.message);
    } else {
      setResetMsg("If that email has an account, a reset link is on its way. Check your inbox.");
    }
  }

  return (
    <div
      style={{
        background: COLORS.bg,
        color: COLORS.ink,
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontFamily: "'Inter', sans-serif",
      }}
    >
      <form
        onSubmit={handleSignIn}
        style={{
          background: COLORS.surface,
          border: `1px solid ${COLORS.hair}`,
          borderRadius: "10px",
          padding: "32px",
          width: "320px",
        }}
      >
        <h1
          style={{
            fontFamily: "'Anton', sans-serif",
            fontSize: "28px",
            margin: "0 0 4px",
            letterSpacing: "1px",
          }}
        >
          THE BIG BOARD
        </h1>
        <p style={{ fontSize: "12.5px", color: COLORS.inkDim, margin: "0 0 24px" }}>
          Sign in with your team account.
        </p>

        <label style={{ fontSize: "11px", color: COLORS.inkDim, display: "block", marginBottom: "4px" }}>
          Email
        </label>
        <input
          type="email"
          required
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          placeholder="name@agency.com"
          style={inputStyle}
        />

        <label style={{ fontSize: "11px", color: COLORS.inkDim, display: "block", margin: "14px 0 4px" }}>
          Password
        </label>
        <input
          type="password"
          required
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          placeholder="••••••••"
          style={inputStyle}
        />

        {error && (
          <div style={{ color: "#D98080", fontSize: "12px", marginTop: "12px" }}>{error}</div>
        )}

        <button
          type="submit"
          disabled={loading}
          style={{
            width: "100%",
            marginTop: "20px",
            padding: "10px",
            borderRadius: "6px",
            border: `1px solid ${COLORS.accent}`,
            background: "rgba(201,138,62,0.12)",
            color: COLORS.accent,
            fontFamily: "'Bebas Neue', sans-serif",
            fontSize: "16px",
            letterSpacing: "1px",
            cursor: "pointer",
          }}
        >
          {loading ? "Signing in…" : "Sign in"}
        </button>

        <div style={{ textAlign: "center", marginTop: "14px" }}>
          <button
            type="button"
            onClick={() => {
              setResetOpen((v) => !v);
              setResetEmail(email);
              setResetMsg("");
            }}
            style={{
              background: "none",
              border: "none",
              color: COLORS.inkDim,
              fontSize: "11.5px",
              cursor: "pointer",
              textDecoration: "underline",
              padding: 0,
            }}
          >
            {resetOpen ? "Hide" : "Forgot or need to set your password?"}
          </button>
        </div>

        {resetOpen && (
          <div style={{ marginTop: "12px", paddingTop: "14px", borderTop: `1px solid ${COLORS.hair}` }}>
            <label style={{ fontSize: "11px", color: COLORS.inkDim, display: "block", marginBottom: "4px" }}>
              Email for reset link
            </label>
            <input
              type="email"
              value={resetEmail}
              onChange={(e) => setResetEmail(e.target.value)}
              placeholder="name@agency.com"
              style={inputStyle}
            />
            <button
              type="button"
              onClick={handleResetPassword}
              disabled={resetLoading}
              style={{
                width: "100%",
                marginTop: "10px",
                padding: "8px",
                borderRadius: "5px",
                border: `1px solid ${COLORS.hair}`,
                background: "transparent",
                color: COLORS.ink,
                fontSize: "12.5px",
                cursor: "pointer",
              }}
            >
              {resetLoading ? "Sending…" : "Send reset link"}
            </button>
            {resetMsg && (
              <div style={{ fontSize: "11.5px", color: COLORS.inkDim, marginTop: "8px", lineHeight: 1.4 }}>
                {resetMsg}
              </div>
            )}
          </div>
        )}

        <p style={{ fontSize: "11px", color: COLORS.inkDim, marginTop: "18px", lineHeight: 1.5 }}>
          No account yet? Ask whoever manages your Supabase project to invite you
          from Authentication → Users. There's no public sign-up on this board.
        </p>
      </form>
    </div>
  );
}

const inputStyle = {
  width: "100%",
  background: "rgba(236,231,220,0.06)",
  border: `1px solid ${COLORS.hair}`,
  color: COLORS.ink,
  borderRadius: "5px",
  padding: "9px 10px",
  fontSize: "13.5px",
  outline: "none",
};
