import { useState, useEffect, useRef, useCallback } from "react";
import axios from "axios";
import * as XLSX from "xlsx";
import {
  Mail, Users, Sliders, History, FileText, Upload, Download, Send,
  CheckCircle2, XCircle, AlertCircle, Play, Pause, Square, Plus,
  Trash2, Edit2, Search, FileSpreadsheet, Eye, Check,
  RefreshCw, Info, ChevronLeft, ChevronRight, Database,
  ArrowRight, LogOut, UserCheck, Lock, Bell, X,
  Shield
} from "lucide-react";

axios.defaults.baseURL = process.env.REACT_APP_API_URL || "http://localhost:5000";

// ─── GOOGLE FONTS LOADER ────────────────────────────────────────────────────
const FontLoader = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,600;1,300;1,400;1,600&family=Instrument+Sans:wght@300;400;500;600&display=swap');
    @import url('https://api.fontshare.com/v2/css?f[]=clash-display@200,300,400,500,600,700&display=swap');
  `}</style>
);

// ─── CSS-IN-JS STYLES ────────────────────────────────────────────────────────
const GlobalStyles = () => (
  <style>{`
    :root {
      --ink: #0d0c0a;
      --cream: #f5f0e8;
      --warm: #f0e8d5;
      --rust: #c94b2a;
      --rust-dark: #a83820;
      --gold: #d4a832;
      --sage: #4a6650;
      --muted: #8a8070;
      --paper: #faf7f2;
      --border: #d8d0c0;
      --panel: #0f1117;
      --panel2: #13181f;
      --glass: rgba(255,255,255,0.03);
      --glow: rgba(201,75,42,0.15);
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Instrument Sans', sans-serif;
      background: var(--paper);
      color: var(--ink);
      overflow-x: hidden;
    }

    /* ── CURSOR ── */
    .sp-cursor {
      position: fixed; width: 10px; height: 10px;
      background: var(--rust); border-radius: 50%;
      pointer-events: none; z-index: 9999;
      transform: translate(-50%,-50%);
      mix-blend-mode: multiply;
      transition: width .15s, height .15s;
    }
    .sp-cursor-ring {
      position: fixed; width: 32px; height: 32px;
      border: 1.5px solid var(--rust); border-radius: 50%;
      pointer-events: none; z-index: 9998;
      transform: translate(-50%,-50%);
      opacity: .45;
    }

    /* ── NOISE OVERLAY ── */
    .sp-noise::after {
      content: '';
      position: fixed; inset: 0;
      background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 200 200' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='.025'/%3E%3C/svg%3E");
      pointer-events: none; z-index: 2000; opacity: .5;
    }

    /* ── LANDING NAV ── */
    .sp-nav {
      position: fixed; top: 0; left: 0; right: 0; z-index: 500;
      display: flex; align-items: center; justify-content: space-between;
      padding: 1.25rem 4vw;
      animation: spFadeDown .8s ease both;
    }
    @keyframes spFadeDown { from{opacity:0;transform:translateY(-12px)} to{opacity:1;transform:none} }

    .sp-logo {
      font-family: 'Clash Display', sans-serif;
      font-weight: 600; font-size: 1.75rem;
      letter-spacing: -.01em; color: var(--ink);
      text-decoration: none; cursor: pointer;
    }
    .sp-logo span { color: var(--rust); }

    .sp-nav-links { display: flex; gap: 2.2rem; list-style: none; align-items: center; }
    .sp-nav-links a, .sp-nav-links button {
      text-decoration: none; background: none; border: none; cursor: pointer;
      font-family: 'Instrument Sans', sans-serif;
      font-size: .78rem; font-weight: 500;
      letter-spacing: .12em; text-transform: uppercase;
      color: var(--ink); opacity: .65; transition: opacity .2s;
    }
    .sp-nav-links a:hover, .sp-nav-links button:hover { opacity: 1; }
    .sp-nav-cta {
      background: var(--ink) !important; color: var(--cream) !important;
      padding: .5rem 1.3rem; border-radius: 2px; opacity: 1 !important;
    }

    /* ── HERO ── */
    .sp-hero {
      min-height: 100vh; display: grid; grid-template-columns: 1fr 1fr;
      overflow: hidden; position: relative;
    }
    .sp-hero-left {
      display: flex; flex-direction: column; justify-content: center;
      padding: 9rem 3rem 5rem 5vw; position: relative; z-index: 2;
    }
    .sp-eyebrow {
      display: flex; align-items: center; gap: .75rem;
      font-family: 'Instrument Sans', sans-serif;
      font-size: .7rem; font-weight: 600; letter-spacing: .3em;
      text-transform: uppercase; color: var(--rust); margin-bottom: 1.8rem;
      opacity: 0; animation: spFadeUp .7s .15s cubic-bezier(.22,1,.36,1) forwards;
    }
    .sp-eyebrow::before {
      content: ''; display: block; width: 32px; height: 1.5px;
      background: var(--rust); flex-shrink: 0;
    }
    .sp-hero-title {
      font-family: 'Clash Display', sans-serif;
      font-weight: 600; font-size: clamp(4.8rem,7.5vw,8.5rem);
      line-height: .9; letter-spacing: -.03em; color: var(--ink);
      text-rendering: optimizeLegibility; -webkit-font-smoothing: antialiased;
    }
    .sp-title-line { display: block; overflow: hidden; padding-bottom: .06em; }
    .sp-title-inner {
      display: block; opacity: 0; transform: translateY(110%);
      transition: opacity .01s, transform .85s cubic-bezier(.16,1,.3,1);
    }
    .sp-title-inner.revealed { opacity: 1; transform: translateY(0); }
    .sp-title-line:nth-child(1) .sp-title-inner { transition-delay: .25s; }
    .sp-title-line:nth-child(2) .sp-title-inner {
      transition-delay: .38s; color: transparent;
      -webkit-text-stroke: 2px var(--ink);
    }
    .sp-title-line:nth-child(3) .sp-title-inner {
      transition-delay: .51s;
      font-family: 'Cormorant Garamond', serif;
      font-style: italic; font-weight: 300; font-size: 1.08em;
      color: var(--rust); letter-spacing: -.01em;
    }
    @keyframes spFadeUp { from{opacity:0;transform:translateY(20px)} to{opacity:1;transform:none} }

    .sp-hero-desc {
      margin-top: 2.2rem;
      font-family: 'Instrument Sans', sans-serif;
      font-size: 1.05rem; line-height: 1.75; color: var(--muted);
      max-width: 400px; font-weight: 300; letter-spacing: .01em;
      opacity: 0; animation: spFadeUp .8s .75s cubic-bezier(.22,1,.36,1) forwards;
    }
    .sp-hero-desc strong { font-weight: 500; color: var(--ink); border-bottom: 1.5px solid var(--rust); }

    .sp-hero-actions {
      margin-top: 2.8rem; display: flex; gap: 1.2rem; align-items: center;
      opacity: 0; animation: spFadeUp .8s .9s cubic-bezier(.22,1,.36,1) forwards;
    }
    .sp-btn-primary {
      background: var(--rust); color: var(--cream); border: none;
      padding: .9rem 2.2rem;
      font-family: 'Instrument Sans', sans-serif;
      font-size: .82rem; font-weight: 600; letter-spacing: .1em;
      cursor: pointer; border-radius: 2px; text-transform: uppercase;
      text-decoration: none; transition: background .25s, transform .15s;
      display: inline-block;
    }
    .sp-btn-primary:hover { background: var(--rust-dark); transform: translateY(-2px); }
    .sp-btn-ghost {
      display: flex; align-items: center; gap: .5rem;
      color: var(--ink); text-decoration: none; cursor: pointer;
      background: none; border: none;
      font-family: 'Instrument Sans', sans-serif;
      font-size: .82rem; font-weight: 500; letter-spacing: .1em; text-transform: uppercase;
      transition: color .2s;
    }
    .sp-btn-ghost::after { content: '→'; font-size: 1rem; transition: transform .25s; }
    .sp-btn-ghost:hover { color: var(--rust); }
    .sp-btn-ghost:hover::after { transform: translateX(4px); }

    .sp-stats {
      margin-top: 4rem; display: flex; gap: 2.8rem;
      opacity: 0; animation: spFadeUp .8s 1.05s cubic-bezier(.22,1,.36,1) forwards;
    }
    .sp-stat-num {
      font-family: 'Clash Display', sans-serif; font-weight: 600;
      font-size: 2.4rem; letter-spacing: -.03em; color: var(--ink); line-height: 1;
    }
    .sp-stat-label {
      font-family: 'Instrument Sans', sans-serif;
      font-size: .68rem; font-weight: 500; text-transform: uppercase;
      letter-spacing: .2em; color: var(--muted); margin-top: .3rem;
    }
    .sp-stat-div { width: 1px; background: var(--border); align-self: stretch; }

    /* ── HERO RIGHT ── */
    .sp-hero-right {
      position: relative; overflow: hidden; background: var(--warm);
    }
    .sp-grid-bg {
      position: absolute; inset: 0;
      background-image: linear-gradient(var(--cream) 1px, transparent 1px),
        linear-gradient(90deg, var(--cream) 1px, transparent 1px);
      background-size: 40px 40px; opacity: .5;
    }
    .sp-envelope-scene { position: absolute; inset: 0; }
    .sp-env-card {
      position: absolute; background: var(--paper);
      border: 1.5px solid #d4cabb; border-radius: 4px;
      display: flex; flex-direction: column; gap: 6px;
      padding: 18px 22px; box-shadow: 0 8px 40px rgba(0,0,0,.06);
      animation: spFloat linear infinite;
    }
    .sp-env-card .env-to { font-size: .68rem; letter-spacing: .1em; text-transform: uppercase; color: var(--muted); }
    .sp-env-card .env-subject { font-family: 'Cormorant Garamond', serif; font-size: 1rem; color: var(--ink); }
    .sp-env-card .env-meta { font-size: .7rem; color: var(--muted); }
    .sp-env-card .env-bar { height: 6px; border-radius: 2px; background: var(--rust); margin-top: 6px; }
    @keyframes spFloat {
      0%{transform:translateY(0)rotate(var(--r))} 50%{transform:translateY(-14px)rotate(var(--r))} 100%{transform:translateY(0)rotate(var(--r))}
    }
    .sp-env-card:nth-child(1){--r:-3deg;width:240px;top:14%;left:8%;animation-duration:6s}
    .sp-env-card:nth-child(2){--r:2deg;width:210px;top:55%;left:5%;animation-duration:7s;animation-delay:1s}
    .sp-env-card:nth-child(3){--r:-1.5deg;width:260px;top:28%;right:4%;animation-duration:8s;animation-delay:.5s}
    .sp-env-card:nth-child(4){--r:3deg;width:200px;bottom:10%;right:8%;animation-duration:5.5s;animation-delay:1.5s}
    .sp-stamp {
      position: absolute; bottom: 2rem; left: 50%; transform: translateX(-50%) rotate(-8deg);
      font-family: 'Clash Display', sans-serif; font-size: 3rem;
      letter-spacing: .15em; color: var(--rust);
      border: 5px solid var(--rust); padding: .4rem 1.2rem;
      opacity: .1; white-space: nowrap;
    }
    .sp-hero-right::after {
      content: ''; position: absolute; top: -60px; left: -60px;
      width: 160%; height: 60px; background: var(--paper);
      transform: rotate(4deg); transform-origin: top left;
    }

    /* ── MARQUEE ── */
    .sp-marquee { background: var(--ink); overflow: hidden; padding: .8rem 0; }
    .sp-marquee-track {
      display: flex; white-space: nowrap;
      animation: spScroll 22s linear infinite;
    }
    @keyframes spScroll { from{transform:translateX(0)} to{transform:translateX(-50%)} }
    .sp-marquee-track span {
      font-family: 'Clash Display', sans-serif; font-weight: 500;
      font-size: .85rem; letter-spacing: .2em;
      color: var(--cream); opacity: .5; padding: 0 2rem;
    }
    .sp-marquee-track .dot { color: var(--rust); opacity: 1 !important; }

    /* ── SECTION COMMON ── */
    .sp-section-label {
      font-family: 'Instrument Sans', sans-serif;
      font-size: .68rem; font-weight: 600; letter-spacing: .28em;
      text-transform: uppercase; color: var(--rust); margin-bottom: .8rem;
    }
    .sp-section-title {
      font-family: 'Clash Display', sans-serif; font-weight: 500;
      font-size: clamp(2.2rem,4vw,3.2rem); line-height: 1.08;
      letter-spacing: -.025em; color: var(--ink);
    }
    .sp-section-title em {
      font-family: 'Cormorant Garamond', serif; font-style: italic;
      font-weight: 300; font-size: 1.05em; color: var(--rust);
    }

    /* ── FEATURES ── */
    .sp-features { padding: 8rem 5vw; }
    .sp-features-grid {
      margin-top: 4rem; display: grid; grid-template-columns: repeat(3,1fr);
      gap: 1px; background: var(--border); border: 1px solid var(--border);
    }
    .sp-feat-card {
      background: var(--paper); padding: 2.8rem 2.4rem;
      transition: background .3s; position: relative; overflow: hidden;
    }
    .sp-feat-card::before {
      content: ''; position: absolute; bottom: 0; left: 0; right: 0;
      height: 3px; background: var(--rust); transform: scaleX(0);
      transform-origin: left; transition: transform .35s;
    }
    .sp-feat-card:hover { background: var(--warm); }
    .sp-feat-card:hover::before { transform: scaleX(1); }
    .sp-feat-num {
      font-family: 'Clash Display', sans-serif; font-size: 3.5rem;
      color: #e8e0d4; position: absolute; top: 1rem; right: 1.5rem; line-height: 1;
    }
    .sp-feat-icon { font-size: 2rem; margin-bottom: 1.2rem; display: block; }
    .sp-feat-title {
      font-family: 'Clash Display', sans-serif; font-weight: 500;
      font-size: 1.25rem; letter-spacing: -.01em; color: var(--ink); margin-bottom: .8rem;
    }
    .sp-feat-desc {
      font-family: 'Instrument Sans', sans-serif;
      font-size: .88rem; line-height: 1.75; color: var(--muted); font-weight: 300;
    }

    /* ── HOW IT WORKS ── */
    .sp-how {
      padding: 8rem 5vw; background: var(--ink); color: var(--cream);
      position: relative; overflow: hidden;
    }
    .sp-how::before {
      content: 'SEND'; position: absolute;
      font-family: 'Clash Display', sans-serif; font-size: 22vw;
      color: white; opacity: .03; top: 50%; left: 50%;
      transform: translate(-50%,-50%); letter-spacing: .15em; white-space: nowrap;
    }
    .sp-steps { margin-top: 4rem; display: grid; grid-template-columns: repeat(4,1fr); position: relative; }
    .sp-steps::before {
      content: ''; position: absolute; top: 1.4rem; left: 2.5rem; right: 2.5rem;
      height: 1px; background: rgba(255,255,255,.1);
    }
    .sp-step { padding: 0 2rem; }
    .sp-step-num {
      width: 44px; height: 44px; border: 1.5px solid var(--gold); border-radius: 50%;
      display: flex; align-items: center; justify-content: center;
      font-family: 'Clash Display', sans-serif; font-size: 1.1rem;
      color: var(--gold); margin-bottom: 1.5rem; background: var(--ink); position: relative; z-index: 1;
    }
    .sp-step-title {
      font-family: 'Clash Display', sans-serif; font-weight: 500; font-size: 1.1rem;
      letter-spacing: -.01em; color: var(--cream); margin-bottom: .6rem;
    }
    .sp-step-desc {
      font-family: 'Instrument Sans', sans-serif;
      font-size: .84rem; line-height: 1.7; color: rgba(245,240,232,.5); font-weight: 300;
    }

    /* ── PRICING ── */
    .sp-pricing { padding: 8rem 5vw; }
    .sp-pricing-grid {
      margin-top: 4rem; display: grid; grid-template-columns: repeat(3,1fr);
      gap: 1.5rem; max-width: 900px;
    }
    .sp-price-card {
      border: 1.5px solid var(--border); border-radius: 4px;
      padding: 2.5rem 2rem; background: var(--paper); position: relative;
      transition: box-shadow .3s, border-color .3s;
    }
    .sp-price-card:hover { box-shadow: 0 12px 50px rgba(0,0,0,.08); border-color: var(--rust); }
    .sp-price-card.featured { background: var(--ink); border-color: var(--ink); }
    .sp-price-badge {
      position: absolute; top: -12px; left: 50%; transform: translateX(-50%);
      background: var(--rust); color: var(--cream); font-size: .65rem;
      letter-spacing: .15em; text-transform: uppercase; padding: 3px 14px;
      border-radius: 20px; white-space: nowrap;
    }
    .sp-price-plan { font-size: .72rem; letter-spacing: .2em; text-transform: uppercase; color: var(--muted); margin-bottom: 1rem; }
    .sp-price-card.featured .sp-price-plan { color: rgba(245,240,232,.5); }
    .sp-price-amount {
      font-family: 'Clash Display', sans-serif; font-weight: 600;
      font-size: 3.5rem; line-height: 1; letter-spacing: -.04em; color: var(--ink);
    }
    .sp-price-card.featured .sp-price-amount { color: var(--cream); }
    .sp-price-period { font-size: .8rem; color: var(--muted); margin-bottom: 1.5rem; }
    .sp-price-card.featured .sp-price-period { color: rgba(245,240,232,.5); }
    .sp-price-features { list-style: none; display: flex; flex-direction: column; gap: .65rem; margin-bottom: 2rem; }
    .sp-price-features li {
      font-size: .85rem; color: var(--muted); padding-left: 1.2rem; position: relative;
    }
    .sp-price-features li::before { content: '—'; position: absolute; left: 0; color: var(--rust); }
    .sp-price-card.featured .sp-price-features li { color: rgba(245,240,232,.7); }
    .sp-btn-plan {
      display: block; text-align: center; text-decoration: none; cursor: pointer;
      width: 100%; padding: .75rem; border-radius: 2px;
      font-size: .8rem; font-weight: 500; letter-spacing: .1em; text-transform: uppercase;
      border: 1.5px solid var(--ink); color: var(--ink); background: transparent;
      transition: background .2s, color .2s;
    }
    .sp-btn-plan:hover { background: var(--ink); color: var(--cream); }
    .sp-price-card.featured .sp-btn-plan {
      background: var(--rust); border-color: var(--rust); color: var(--cream);
    }
    .sp-price-card.featured .sp-btn-plan:hover { background: var(--rust-dark); }

    /* ── TESTIMONIALS ── */
    .sp-testimonials { padding: 8rem 5vw; background: var(--warm); }
    .sp-testi-grid { margin-top: 3.5rem; display: grid; grid-template-columns: repeat(3,1fr); gap: 1.5rem; }
    .sp-testi-card {
      background: var(--paper); border: 1.5px solid var(--border);
      border-radius: 4px; padding: 2rem;
    }
    .sp-testi-quote {
      font-family: 'Cormorant Garamond', serif; font-size: 1.2rem;
      font-weight: 400; line-height: 1.6; color: var(--ink); margin-bottom: 1.5rem; font-style: italic;
    }
    .sp-testi-author { font-family: 'Instrument Sans', sans-serif; font-size: .78rem; font-weight: 600; color: var(--ink); }
    .sp-testi-role { font-family: 'Instrument Sans', sans-serif; font-size: .72rem; color: var(--muted); margin-top: 2px; }

    /* ── CTA ── */
    .sp-cta {
      padding: 8rem 5vw; display: flex; flex-direction: column;
      align-items: center; text-align: center; position: relative;
    }
    .sp-cta::before {
      content: ''; position: absolute; width: 600px; height: 600px;
      background: radial-gradient(circle, rgba(201,75,42,.07) 0%, transparent 70%);
      top: 50%; left: 50%; transform: translate(-50%,-50%);
    }
    .sp-cta-title {
      font-family: 'Clash Display', sans-serif; font-weight: 600;
      font-size: clamp(3.5rem,6vw,6rem); letter-spacing: -.03em;
      line-height: .95; color: var(--ink); max-width: 700px;
    }
    .sp-cta-title span { color: var(--rust); }
    .sp-cta-sub { margin-top: 1.5rem; font-size: 1rem; color: var(--muted); max-width: 400px; line-height: 1.7; font-weight: 300; }
    .sp-email-wrap { display: flex; border: 1.5px solid var(--border); border-radius: 2px; overflow: hidden; background: var(--paper); margin-top: 2.5rem; }
    .sp-email-input { padding: .8rem 1.2rem; border: none; outline: none; font-family: 'Instrument Sans', sans-serif; font-size: .88rem; background: transparent; color: var(--ink); width: 240px; }
    .sp-email-input::placeholder { color: var(--muted); }
    .sp-email-btn {
      background: var(--rust); color: var(--cream); border: none; cursor: pointer;
      padding: .8rem 1.4rem; font-family: 'Instrument Sans', sans-serif;
      font-size: .82rem; font-weight: 500; letter-spacing: .08em; text-transform: uppercase;
      transition: background .2s;
    }
    .sp-email-btn:hover { background: var(--rust-dark); }

    /* ── FOOTER ── */
    .sp-footer {
      background: var(--ink); color: var(--cream); padding: 3.5rem 5vw;
      display: flex; align-items: center; justify-content: space-between;
    }
    .sp-footer-logo { font-family: 'Clash Display', sans-serif; font-weight: 600; font-size: 1.55rem; letter-spacing: -.01em; }
    .sp-footer-logo span { color: var(--rust); }
    .sp-footer-links { display: flex; gap: 2rem; list-style: none; }
    .sp-footer-links a { text-decoration: none; font-size: .75rem; color: rgba(245,240,232,.45); letter-spacing: .1em; text-transform: uppercase; transition: color .2s; }
    .sp-footer-links a:hover { color: var(--cream); }
    .sp-footer-copy { font-size: .75rem; color: rgba(245,240,232,.35); }

    /* ── AUTH SCREEN ── */
    .sp-auth {
      min-height: 100vh; display: flex; align-items: center; justify-content: center;
      background: var(--paper); position: relative;
    }
    .sp-auth-card {
      background: white; border: 1.5px solid var(--border); border-radius: 4px;
      padding: 3rem 2.5rem; width: 100%; max-width: 420px;
      box-shadow: 0 24px 80px rgba(0,0,0,.08);
      animation: spFadeUp .7s ease both;
    }
    .sp-auth-title {
      font-family: 'Clash Display', sans-serif; font-weight: 600;
      font-size: 1.9rem; letter-spacing: -.02em; color: var(--ink); text-align: center;
    }
    .sp-auth-title span { color: var(--rust); }
    .sp-auth-subtitle { font-family: 'Instrument Sans', sans-serif; font-size: .72rem; font-weight: 500; letter-spacing: .2em; text-transform: uppercase; color: var(--muted); text-align: center; margin-top: .3rem; }
    .sp-auth-tab {
      flex: 1; padding: .75rem; background: transparent; border: none; cursor: pointer;
      font-family: 'Instrument Sans', sans-serif; font-size: .78rem; font-weight: 600;
      letter-spacing: .1em; text-transform: uppercase; color: var(--muted);
      border-bottom: 2px solid transparent; transition: all .2s;
    }
    .sp-auth-tab.active { color: var(--rust); border-bottom-color: var(--rust); }
    .sp-auth-label {
      display: block; font-family: 'Instrument Sans', sans-serif;
      font-size: .68rem; font-weight: 600; letter-spacing: .15em;
      text-transform: uppercase; color: var(--muted); margin-bottom: .5rem;
    }
    .sp-auth-input {
      width: 100%; padding: .75rem 1rem .75rem 2.8rem;
      border: 1.5px solid var(--border); border-radius: 2px;
      font-family: 'Instrument Sans', sans-serif; font-size: .88rem;
      background: var(--paper); color: var(--ink); outline: none;
      transition: border-color .2s;
    }
    .sp-auth-input:focus { border-color: var(--rust); }
    .sp-auth-submit {
      width: 100%; padding: .9rem; background: var(--rust); color: var(--cream); border: none; cursor: pointer;
      font-family: 'Instrument Sans', sans-serif; font-size: .82rem; font-weight: 600;
      letter-spacing: .1em; text-transform: uppercase; border-radius: 2px;
      transition: background .2s, transform .15s; display: flex; align-items: center; justify-content: center; gap: .5rem;
    }
    .sp-auth-submit:hover { background: var(--rust-dark); transform: translateY(-1px); }
    .sp-auth-submit:disabled { opacity: .6; cursor: not-allowed; transform: none; }

    /* ── DASHBOARD ── */
    .sp-dash {
      display: flex; min-height: 100vh;
      background: #0c0e12; color: var(--cream);
      font-family: 'Instrument Sans', sans-serif;
    }
    .sp-sidebar {
      width: 260px; flex-shrink: 0; display: flex; flex-direction: column;
      background: #0f1115; border-right: 1px solid rgba(245,240,232,.06);
      position: relative;
    }
    .sp-sidebar::before {
      content: ''; position: absolute; top: 0; left: 0; right: 0; height: 1px;
      background: linear-gradient(90deg, transparent, var(--rust), transparent); opacity: .4;
    }
    .sp-sidebar-logo {
      padding: 1.75rem 1.5rem 1.25rem;
      border-bottom: 1px solid rgba(245,240,232,.05);
    }
    .sp-sidebar-logo .logo-mark {
      font-family: 'Clash Display', sans-serif; font-weight: 600;
      font-size: 1.5rem; letter-spacing: -.01em;
    }
    .sp-sidebar-logo .logo-mark span { color: var(--rust); }
    .sp-sidebar-logo .logo-tag {
      font-size: .62rem; font-weight: 600; letter-spacing: .2em;
      text-transform: uppercase; color: rgba(245,240,232,.3); margin-top: .15rem;
    }
    .sp-sidebar-nav { flex: 1; padding: 1.25rem .75rem; }
    .sp-nav-item {
      width: 100%; display: flex; align-items: center; justify-content: space-between;
      gap: .75rem; padding: .75rem 1rem; border-radius: 2px; border: none; cursor: pointer;
      font-family: 'Instrument Sans', sans-serif; font-size: .82rem; font-weight: 500;
      letter-spacing: .02em; background: transparent; color: rgba(245,240,232,.45);
      transition: all .2s; margin-bottom: 2px; text-align: left;
    }
    .sp-nav-item:hover { color: var(--cream); background: rgba(245,240,232,.04); }
    .sp-nav-item.active {
      background: rgba(201,75,42,.1); color: var(--cream);
      border-left: 2px solid var(--rust); padding-left: calc(1rem - 2px);
    }
    .sp-nav-item .nav-left { display: flex; align-items: center; gap: .75rem; }
    .sp-badge {
      font-size: .62rem; font-weight: 700; padding: 2px 8px;
      border-radius: 20px; background: rgba(201,75,42,.15);
      color: var(--rust); border: 1px solid rgba(201,75,42,.2);
    }
    .sp-sidebar-footer { padding: 1rem .75rem 1.5rem; border-top: 1px solid rgba(245,240,232,.05); }
    .sp-status-dot {
      width: 8px; height: 8px; border-radius: 50%;
      background: #4ade80; box-shadow: 0 0 0 3px rgba(74,222,128,.15);
    }
    .sp-status-dot.offline { background: #f87171; box-shadow: 0 0 0 3px rgba(248,113,113,.15); }
    .sp-status-dot.checking { background: #94a3b8; }

    /* ── DASH MAIN ── */
    .sp-main { flex: 1; display: flex; flex-direction: column; height: 100vh; overflow: hidden; }
    .sp-topbar {
      height: 60px; padding: 0 2rem; display: flex; align-items: center;
      justify-content: space-between; border-bottom: 1px solid rgba(245,240,232,.06);
      background: #0c0e12; flex-shrink: 0;
    }
    .sp-topbar-title {
      font-family: 'Clash Display', sans-serif; font-weight: 500;
      font-size: 1.1rem; letter-spacing: -.01em; color: var(--cream);
    }
    .sp-content { flex: 1; overflow-y: auto; padding: 2rem; background: #080a0e; }

    /* ── DASH CARDS ── */
    .sp-card {
      background: #0f1115; border: 1px solid rgba(245,240,232,.06);
      border-radius: 4px; padding: 1.75rem;
      transition: border-color .2s;
    }
    .sp-card:hover { border-color: rgba(201,75,42,.25); }
    .sp-card-title {
      font-family: 'Clash Display', sans-serif; font-weight: 500;
      font-size: 1rem; letter-spacing: -.01em; color: var(--cream);
      display: flex; align-items: center; gap: .6rem; margin-bottom: 1.25rem;
    }
    .sp-label {
      font-size: .65rem; font-weight: 600; letter-spacing: .2em;
      text-transform: uppercase; color: rgba(245,240,232,.35); margin-bottom: .4rem;
    }
    .sp-input {
      width: 100%; padding: .7rem .9rem;
      background: #13181f; border: 1px solid rgba(245,240,232,.08);
      border-radius: 2px; color: var(--cream); outline: none;
      font-family: 'Instrument Sans', sans-serif; font-size: .85rem;
      transition: border-color .2s;
    }
    .sp-input:focus { border-color: var(--rust); }
    .sp-input::placeholder { color: rgba(245,240,232,.2); }
    .sp-btn-dash {
      padding: .65rem 1.25rem; border-radius: 2px; border: none; cursor: pointer;
      font-family: 'Instrument Sans', sans-serif; font-size: .78rem;
      font-weight: 600; letter-spacing: .08em; text-transform: uppercase;
      transition: all .2s; display: inline-flex; align-items: center; gap: .4rem;
    }
    .sp-btn-rust { background: var(--rust); color: var(--cream); }
    .sp-btn-rust:hover { background: var(--rust-dark); transform: translateY(-1px); }
    .sp-btn-ink { background: rgba(245,240,232,.06); color: rgba(245,240,232,.7); border: 1px solid rgba(245,240,232,.08); }
    .sp-btn-ink:hover { background: rgba(245,240,232,.1); color: var(--cream); }
    .sp-btn-danger { background: rgba(239,68,68,.1); color: #f87171; border: 1px solid rgba(239,68,68,.2); }
    .sp-btn-danger:hover { background: rgba(239,68,68,.2); }
    .sp-btn-disabled { opacity: .45; cursor: not-allowed !important; transform: none !important; }

    .sp-stat-card {
      background: #0f1115; border: 1px solid rgba(245,240,232,.06);
      border-radius: 4px; padding: 1.5rem;
      display: flex; align-items: center; justify-content: space-between;
      position: relative; overflow: hidden;
    }
    .sp-stat-card::before {
      content: ''; position: absolute; bottom: 0; left: 0; right: 0;
      height: 2px; background: var(--rust); transform: scaleX(0);
      transform-origin: left; transition: transform .4s ease;
    }
    .sp-stat-card:hover::before { transform: scaleX(1); }
    .sp-stat-big {
      font-family: 'Clash Display', sans-serif; font-weight: 600;
      font-size: 2.5rem; letter-spacing: -.04em; line-height: 1;
    }

    /* ── TAGS / CHIPS ── */
    .sp-chip {
      display: inline-flex; align-items: center; gap: .3rem;
      padding: 3px 10px; border-radius: 20px;
      font-size: .65rem; font-weight: 700; letter-spacing: .1em; text-transform: uppercase;
    }
    .sp-chip-green { background: rgba(74,222,128,.1); color: #4ade80; border: 1px solid rgba(74,222,128,.2); }
    .sp-chip-red { background: rgba(248,113,113,.1); color: #f87171; border: 1px solid rgba(248,113,113,.2); }
    .sp-chip-amber { background: rgba(251,191,36,.1); color: #fbbf24; border: 1px solid rgba(251,191,36,.2); }

    /* ── TABLE ── */
    .sp-table { width: 100%; border-collapse: collapse; font-size: .82rem; }
    .sp-table th {
      padding: .7rem 1rem; text-align: left; font-size: .62rem; font-weight: 700;
      letter-spacing: .15em; text-transform: uppercase; color: rgba(245,240,232,.3);
      border-bottom: 1px solid rgba(245,240,232,.05); background: rgba(245,240,232,.02);
    }
    .sp-table td {
      padding: .8rem 1rem; border-bottom: 1px solid rgba(245,240,232,.04);
      color: rgba(245,240,232,.75); vertical-align: middle;
    }
    .sp-table tr:hover td { background: rgba(245,240,232,.02); }
    .sp-table tr:last-child td { border-bottom: none; }

    /* ── TEXTAREA / CODE ── */
    .sp-textarea {
      resize: none; font-family: 'JetBrains Mono', 'Fira Code', monospace;
      font-size: .8rem; line-height: 1.6;
    }

    /* ── PROGRESS BAR ── */
    .sp-progress-bar {
      height: 3px; background: rgba(245,240,232,.06); border-radius: 20px; overflow: hidden;
    }
    .sp-progress-fill {
      height: 100%; background: linear-gradient(90deg, var(--rust), var(--gold));
      border-radius: 20px; transition: width .4s ease;
    }

    /* ── WEB ALERTS ── */
    .sp-alerts {
      position: fixed; top: 1.25rem; right: 1.25rem; z-index: 9000;
      display: flex; flex-direction: column; gap: .6rem; max-width: 380px; width: 100%;
    }
    .sp-alert {
      display: flex; align-items: flex-start; gap: .85rem;
      padding: 1rem 1.1rem; border-radius: 3px;
      border-left: 3px solid;
      background: #0f1115; box-shadow: 0 8px 40px rgba(0,0,0,.4);
      animation: spAlertIn .4s cubic-bezier(.16,1,.3,1) both;
      backdrop-filter: blur(12px);
    }
    .sp-alert.leaving { animation: spAlertOut .35s ease forwards; }
    @keyframes spAlertIn {
      from { opacity: 0; transform: translateX(24px) scale(.97); }
      to   { opacity: 1; transform: translateX(0) scale(1); }
    }
    @keyframes spAlertOut {
      from { opacity: 1; transform: translateX(0) scale(1); max-height: 100px; margin-bottom: 0; }
      to   { opacity: 0; transform: translateX(24px) scale(.95); max-height: 0; margin-bottom: -0.6rem; }
    }
    .sp-alert.success { border-color: #4ade80; }
    .sp-alert.error   { border-color: #f87171; }
    .sp-alert.warning { border-color: #fbbf24; }
    .sp-alert.info    { border-color: var(--rust); }
    .sp-alert-icon { flex-shrink: 0; margin-top: 1px; }
    .sp-alert.success .sp-alert-icon { color: #4ade80; }
    .sp-alert.error   .sp-alert-icon { color: #f87171; }
    .sp-alert.warning .sp-alert-icon { color: #fbbf24; }
    .sp-alert.info    .sp-alert-icon { color: var(--rust); }
    .sp-alert-title {
      font-family: 'Clash Display', sans-serif; font-weight: 500;
      font-size: .92rem; letter-spacing: -.01em; color: var(--cream); line-height: 1.1;
    }
    .sp-alert-msg { font-size: .78rem; color: rgba(245,240,232,.55); line-height: 1.5; margin-top: .2rem; }
    .sp-alert-close {
      margin-left: auto; flex-shrink: 0; background: none; border: none; cursor: pointer;
      color: rgba(245,240,232,.3); padding: 2px; transition: color .2s;
    }
    .sp-alert-close:hover { color: var(--cream); }
    .sp-alert-progress {
      position: absolute; bottom: 0; left: 0; height: 2px;
      background: currentColor; border-radius: 0 0 0 3px;
      animation: spAlertTimer linear forwards;
    }
    .sp-alert { position: relative; }

    /* ── MODAL ── */
    .sp-modal-bg {
      position: fixed; inset: 0; background: rgba(0,0,0,.7);
      backdrop-filter: blur(6px); z-index: 8000;
      display: flex; align-items: center; justify-content: center; padding: 1rem;
      animation: spFadeIn .2s ease;
    }
    @keyframes spFadeIn { from{opacity:0} to{opacity:1} }
    .sp-modal {
      background: #0f1115; border: 1px solid rgba(245,240,232,.08);
      border-radius: 4px; padding: 2rem; width: 100%; max-width: 440px;
      box-shadow: 0 40px 120px rgba(0,0,0,.6);
      animation: spFadeUp .35s cubic-bezier(.16,1,.3,1) both;
    }
    .sp-modal-title {
      font-family: 'Clash Display', sans-serif; font-weight: 500;
      font-size: 1.25rem; letter-spacing: -.01em; color: var(--cream); margin-bottom: .4rem;
    }
    .sp-modal-sub { font-size: .8rem; color: rgba(245,240,232,.4); margin-bottom: 1.5rem; }

    /* ── CAMPAIGN OVERLAY ── */
    .sp-overlay {
      position: fixed; inset: 0; background: rgba(0,0,0,.88);
      backdrop-filter: blur(10px); z-index: 8500;
      display: flex; align-items: center; justify-content: center; padding: 1rem;
    }
    .sp-overlay-card {
      background: #0f1115; border: 1px solid rgba(201,75,42,.2);
      border-radius: 4px; padding: 2.5rem; width: 100%; max-width: 640px;
      box-shadow: 0 0 80px rgba(201,75,42,.08), 0 40px 120px rgba(0,0,0,.6);
      animation: spFadeUp .4s cubic-bezier(.16,1,.3,1) both;
    }
    .sp-console {
      background: #080a0e; border: 1px solid rgba(245,240,232,.05);
      border-radius: 2px; padding: 1rem; height: 180px; overflow-y: auto;
      font-family: 'JetBrains Mono', 'Fira Code', monospace; font-size: .75rem;
      line-height: 1.7;
    }
    .sp-console::-webkit-scrollbar { width: 4px; }
    .sp-console::-webkit-scrollbar-thumb { background: rgba(245,240,232,.1); border-radius: 2px; }

    /* ── DRAWER ── */
    .sp-drawer-bg {
      position: fixed; inset: 0; background: rgba(0,0,0,.6);
      backdrop-filter: blur(6px); z-index: 8000;
    }
    .sp-drawer {
      position: fixed; top: 0; right: 0; bottom: 0; width: 100%; max-width: 500px;
      background: #0f1115; border-left: 1px solid rgba(245,240,232,.06);
      z-index: 8001; display: flex; flex-direction: column;
      animation: spDrawerIn .35s cubic-bezier(.16,1,.3,1) both;
      box-shadow: -40px 0 100px rgba(0,0,0,.5);
    }
    @keyframes spDrawerIn { from{transform:translateX(100%)} to{transform:translateX(0)} }

    /* ── FORMAT TOOLBAR ── */
    .sp-toolbar {
      display: flex; gap: 2px; flex-wrap: wrap; padding: .5rem;
      background: #13181f; border: 1px solid rgba(245,240,232,.06);
      border-bottom: none; border-radius: 2px 2px 0 0;
    }
    .sp-toolbar-btn {
      padding: .3rem .6rem; border-radius: 2px; border: none; cursor: pointer;
      font-size: .75rem; font-weight: 600; background: transparent;
      color: rgba(245,240,232,.45); transition: all .15s;
    }
    .sp-toolbar-btn:hover { background: rgba(201,75,42,.15); color: var(--rust); }

    /* ── PLACEHOLDER CHIPS ── */
    .sp-placeholder {
      display: inline-flex; align-items: center; gap: .3rem;
      padding: .2rem .65rem; border-radius: 2px; cursor: pointer;
      font-size: .72rem; font-weight: 600; letter-spacing: .05em;
      background: rgba(201,75,42,.08); color: var(--rust);
      border: 1px solid rgba(201,75,42,.2); transition: all .15s;
    }
    .sp-placeholder:hover { background: rgba(201,75,42,.18); }

    /* ── PREVIEW BROWSER ── */
    .sp-browser {
      border: 1px solid rgba(245,240,232,.08); border-radius: 3px;
      overflow: hidden; display: flex; flex-direction: column;
    }
    .sp-browser-bar {
      background: #13181f; padding: .7rem 1rem;
      display: flex; align-items: center; gap: 1rem;
      border-bottom: 1px solid rgba(245,240,232,.06);
    }
    .sp-browser-dots { display: flex; gap: .35rem; }
    .sp-browser-dot { width: 10px; height: 10px; border-radius: 50%; }
    .sp-browser-body { flex: 1; background: white; padding: 1.5rem; overflow-y: auto; }

    /* ── UPLOAD DROP ZONE ── */
    .sp-dropzone {
      border: 2px dashed rgba(245,240,232,.1); border-radius: 3px;
      padding: 2.5rem; text-align: center; position: relative;
      transition: border-color .2s, background .2s; cursor: pointer;
    }
    .sp-dropzone:hover { border-color: rgba(201,75,42,.4); background: rgba(201,75,42,.04); }
    .sp-dropzone input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }

    /* ── PAGINATION ── */
    .sp-page-btn {
      width: 32px; height: 32px; border-radius: 2px; border: 1px solid rgba(245,240,232,.08);
      background: transparent; color: rgba(245,240,232,.4); cursor: pointer;
      display: flex; align-items: center; justify-content: center; transition: all .15s;
    }
    .sp-page-btn:hover:not(:disabled) { border-color: var(--rust); color: var(--rust); }
    .sp-page-btn:disabled { opacity: .3; cursor: not-allowed; }

    /* ── SLIDER ── */
    input[type=range] { accent-color: var(--rust); }

    /* ── ANIMATIONS ── */
    .sp-fade-in { animation: spFadeUp .5s ease both; }
    .sp-stagger > * { animation: spFadeUp .5s ease both; }
    .sp-stagger > *:nth-child(1){animation-delay:.05s}
    .sp-stagger > *:nth-child(2){animation-delay:.1s}
    .sp-stagger > *:nth-child(3){animation-delay:.15s}
    .sp-stagger > *:nth-child(4){animation-delay:.2s}
    .sp-stagger > *:nth-child(5){animation-delay:.25s}

    .sp-spin { animation: spSpin 1s linear infinite; }
    @keyframes spSpin { to { transform: rotate(360deg); } }

    .sp-pulse { animation: spPulse 2s ease-in-out infinite; }
    @keyframes spPulse { 0%,100%{opacity:1} 50%{opacity:.5} }

    /* ── SCROLLBAR ── */
    .sp-content::-webkit-scrollbar { width: 4px; }
    .sp-content::-webkit-scrollbar-thumb { background: rgba(245,240,232,.08); border-radius: 2px; }

    /* ── TEMPLATE CARD ── */
    .sp-tpl-card {
      background: #0f1115; border: 1px solid rgba(245,240,232,.06);
      border-radius: 4px; padding: 1.5rem; display: flex; flex-direction: column;
      justify-content: space-between; height: 220px;
      transition: border-color .2s, box-shadow .2s;
    }
    .sp-tpl-card:hover { border-color: rgba(201,75,42,.3); box-shadow: 0 8px 40px rgba(201,75,42,.06); }

    /* ── HOVER LIFT ── */
    .sp-lift { transition: transform .2s, box-shadow .2s; }
    .sp-lift:hover { transform: translateY(-2px); box-shadow: 0 12px 40px rgba(0,0,0,.3); }
  `}</style>
);

// ─── ALERT SYSTEM ─────────────────────────────────────────────────────────────
let _alertId = 0;
let _setAlerts = () => {};
export const toast = {
  success: (title, msg) => _addAlert("success", title, msg),
  error:   (title, msg) => _addAlert("error",   title, msg),
  warning: (title, msg) => _addAlert("warning", title, msg),
  info:    (title, msg) => _addAlert("info",     title, msg),
};
function _addAlert(type, title, msg) {
  const id = ++_alertId;
  _setAlerts(prev => [...prev, { id, type, title, msg, leaving: false }]);
  setTimeout(() => _dismissAlert(id), 4500);
}
function _dismissAlert(id) {
  _setAlerts(prev => prev.map(a => a.id === id ? { ...a, leaving: true } : a));
  setTimeout(() => _setAlerts(prev => prev.filter(a => a.id !== id)), 380);
}

const ALERT_ICONS = {
  success: <CheckCircle2 size={16} />,
  error:   <XCircle     size={16} />,
  warning: <AlertCircle size={16} />,
  info:    <Bell        size={16} />,
};

function AlertSystem() {
  const [alerts, setAlerts] = useState([]);
  useEffect(() => { _setAlerts = setAlerts; }, []);
  return (
    <div className="sp-alerts">
      {alerts.map(a => (
        <div key={a.id} className={`sp-alert ${a.type} ${a.leaving ? "leaving" : ""}`}>
          <span className="sp-alert-icon">{ALERT_ICONS[a.type]}</span>
          <div style={{ flex: 1 }}>
            <div className="sp-alert-title">{a.title}</div>
            {a.msg && <div className="sp-alert-msg">{a.msg}</div>}
          </div>
          <button className="sp-alert-close" onClick={() => _dismissAlert(a.id)}>
            <X size={14} />
          </button>
        </div>
      ))}
    </div>
  );
}

// ─── CONFIRM DIALOG ───────────────────────────────────────────────────────────
function ConfirmDialog({ msg, onConfirm, onCancel }) {
  return (
    <div className="sp-modal-bg" onClick={onCancel}>
      <div className="sp-modal" onClick={e => e.stopPropagation()} style={{ maxWidth: 380 }}>
        <div className="sp-modal-title">Are you sure?</div>
        <div className="sp-modal-sub">{msg}</div>
        <div style={{ display: "flex", gap: ".75rem", justifyContent: "flex-end" }}>
          <button className="sp-btn-dash sp-btn-ink" onClick={onCancel}>Cancel</button>
          <button className="sp-btn-dash sp-btn-danger" onClick={onConfirm}>Confirm</button>
        </div>
      </div>
    </div>
  );
}

// ─── MAIN COMPONENT ───────────────────────────────────────────────────────────
export default function BulkMail() {
  // ── Landing / Auth
  const [showLanding, setShowLanding] = useState(true);
  const [ctaEmail, setCtaEmail] = useState("");
  const [revealed, setRevealed] = useState(false);
  const [isAuth, setIsAuth] = useState(false);
  const [authTab, setAuthTab] = useState("login");
  const [authEmail, setAuthEmail] = useState("");
  const [authPass, setAuthPass] = useState("");
  const [authConfirm, setAuthConfirm] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [showAuthPass, setShowAuthPass] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [authToken, setAuthToken] = useState(localStorage.getItem("bulkmail_auth_token") || "");

  // ── Dashboard
  const [activeTab, setActiveTab] = useState("composer");
  const [serverStatus, setServerStatus] = useState("online");

  // ── Composer
  const [subject, setSubject] = useState("");
  const [body, setBody] = useState("");
  const [htmlMode, setHtmlMode] = useState(true);
  const [activeField, setActiveField] = useState("body");
  const [throttle, setThrottle] = useState(1000);
  const [templateName, setTemplateName] = useState("");
  const [showSaveTplModal, setShowSaveTplModal] = useState(false);
  const bodyRef = useRef(null);

  // ── Contacts
  const [contacts, setContacts] = useState([]);
  const [fileHeaders, setFileHeaders] = useState([]);
  const [cSearch, setCSearch] = useState("");
  const [cFilter, setCFilter] = useState("all");
  const [cPage, setCPage] = useState(1);
  const PER_PAGE = 10;
  const [showAddForm, setShowAddForm] = useState(false);
  const [newEmail, setNewEmail] = useState("");
  const [newFields, setNewFields] = useState({});
  const [editingId, setEditingId] = useState(null);
  const [editEmail, setEditEmail] = useState("");
  const [editFields, setEditFields] = useState({});

  // ── SMTP
  const [smtp, setSmtp] = useState({ host: "smtp.gmail.com", port: 587, secure: false, user: "", pass: "", fromName: "", fromEmail: "" });
  const [showPass, setShowPass] = useState(false);
  const [testEmail, setTestEmail] = useState("");
  const [testLoading, setTestLoading] = useState(false);
  const [testResult, setTestResult] = useState(null);

  // ── Templates / Logs
  const [templates, setTemplates] = useState([]);
  const [logs, setLogs] = useState([]);
  const [selectedLog, setSelectedLog] = useState(null);

  // ── Campaign
  const [sending, setSending] = useState({ active: false, progress: 0, total: 0, success: 0, fail: 0, paused: false });
  const [consoleLogs, setConsoleLogs] = useState([]);
  const campaignRef = useRef(false);
  const pauseRef = useRef(false);
  const idxRef = useRef(0);
  const consoleEnd = useRef(null);

  // ── Confirm dialog
  const [confirm, setConfirm] = useState(null);

  // ── Auth Token Synchronization
  useEffect(() => {
    if (authToken) {
      axios.defaults.headers.common["Authorization"] = `Bearer ${authToken}`;
      localStorage.setItem("bulkmail_auth_token", authToken);
    } else {
      delete axios.defaults.headers.common["Authorization"];
      localStorage.removeItem("bulkmail_auth_token");
    }
  }, [authToken]);

  // ── Fetch functions
  const fetchSmtp = useCallback(async () => {
    try {
      const res = await axios.get("/api/smtp");
      if (res.data.success && res.data.data) {
        setSmtp(res.data.data);
      }
    } catch (err) {
      console.error("Failed to fetch SMTP settings:", err.message);
    }
  }, []);

  const fetchTemplates = useCallback(async () => {
    try {
      const res = await axios.get("/api/templates");
      if (res.data.success && res.data.data) {
        setTemplates(res.data.data);
      }
    } catch (err) {
      console.error("Failed to fetch templates:", err.message);
    }
  }, []);

  const fetchContacts = useCallback(async () => {
    try {
      const res = await axios.get("/api/contacts");
      if (res.data.success && res.data.data) {
        const normalized = res.data.data.map(c => ({
          ...c,
          id: c._id || c.id
        }));
        setContacts(normalized);
        if (normalized.length > 0) {
          const first = normalized[0];
          const headers = Object.keys(first.placeholders || {});
          if (!headers.includes("Email")) headers.unshift("Email");
          setFileHeaders(headers);
        }
      }
    } catch (err) {
      console.error("Failed to fetch contacts:", err.message);
    }
  }, []);

  const fetchLogs = useCallback(async () => {
    try {
      const res = await axios.get("/api/logs");
      if (res.data.success && res.data.data) {
        setLogs(res.data.data);
      }
    } catch (err) {
      console.error("Failed to fetch campaign logs:", err.message);
    }
  }, []);

  // ── Verify session on mount
  useEffect(() => {
    const verifySession = async () => {
      if (!authToken) {
        setIsAuth(false);
        setShowLanding(true);
        return;
      }
      try {
        const res = await axios.get("/api/auth/me");
        if (res.data.success) {
          setCurrentUser(res.data.user);
          setIsAuth(true);
          setShowLanding(false);
        } else {
          handleLogout();
        }
      } catch (err) {
        console.error("Session verification failed:", err.message);
        handleLogout();
      }
    };
    verifySession();
  }, [authToken]);

  // ── Fetch dashboard data when authenticated
  useEffect(() => {
    if (isAuth) {
      fetchSmtp();
      fetchTemplates();
      fetchContacts();
      fetchLogs();
    }
  }, [isAuth, fetchSmtp, fetchTemplates, fetchContacts, fetchLogs]);

  // ── Cursor
  const cursorRef = useRef(null);
  const ringRef = useRef(null);
  const mouseRef = useRef({ x: 0, y: 0 });

  // ── Cursor logic (landing only)
  useEffect(() => {
    if (isAuth) return;
    const move = e => {
      mouseRef.current = { x: e.clientX, y: e.clientY };
      if (cursorRef.current) {
        cursorRef.current.style.left = e.clientX + "px";
        cursorRef.current.style.top  = e.clientY + "px";
      }
    };
    window.addEventListener("mousemove", move);
    return () => window.removeEventListener("mousemove", move);
  }, [isAuth]);

  useEffect(() => {
    if (isAuth) return;
    let rx = 0, ry = 0, frame;
    const anim = () => {
      rx += (mouseRef.current.x - rx) * .12;
      ry += (mouseRef.current.y - ry) * .12;
      if (ringRef.current) {
        ringRef.current.style.left = rx + "px";
        ringRef.current.style.top  = ry + "px";
      }
      frame = requestAnimationFrame(anim);
    };
    frame = requestAnimationFrame(anim);
    return () => cancelAnimationFrame(frame);
  }, [isAuth]);

  // ── Title reveal
  useEffect(() => {
    if (showLanding) {
      const t = setTimeout(() => setRevealed(true), 80);
      return () => clearTimeout(t);
    }
  }, [showLanding]);

  // ── Intersection observer for landing cards
  useEffect(() => {
    if (!showLanding || isAuth) return;
    const io = new IntersectionObserver(entries => {
      entries.forEach(e => {
        if (e.isIntersecting) {
          e.target.style.opacity = "1";
          e.target.style.transform = "translateY(0)";
        }
      });
    }, { threshold: .12 });
    const targets = document.querySelectorAll(".sp-feat-card, .sp-step, .sp-price-card, .sp-testi-card");
    targets.forEach(el => {
      el.style.opacity = "0";
      el.style.transform = "translateY(28px)";
      el.style.transition = "opacity .6s ease, transform .6s ease";
      io.observe(el);
    });
    return () => targets.forEach(el => io.unobserve(el));
  }, [showLanding, isAuth]);

  useEffect(() => {
    if (consoleEnd.current) consoleEnd.current.scrollIntoView({ behavior: "smooth" });
  }, [consoleLogs]);

  const logMsg = msg => setConsoleLogs(p => [...p, `[${new Date().toLocaleTimeString()}] ${msg}`]);

  const isValidEmail = str => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(str).trim());

  // ── Auth
  const handleLogin = async (e) => {
    e.preventDefault();
    if (!authEmail.trim() || !authPass.trim()) {
      toast.error("Missing fields", "Please fill in all credentials.");
      return;
    }
    setAuthLoading(true);
    try {
      const res = await axios.post("/api/auth/login", {
        email: authEmail.trim(),
        password: authPass.trim()
      });
      if (res.data.success) {
        setAuthToken(res.data.token);
        setCurrentUser(res.data.user);
        setIsAuth(true);
        setShowLanding(false);
        setAuthEmail("");
        setAuthPass("");
        toast.success("Welcome back!", `Signed in as ${res.data.user.email}`);
      }
    } catch (err) {
      toast.error("Login failed", err.response?.data?.error || "Invalid email or password.");
    } finally {
      setAuthLoading(false);
    }
  };

  const handleSignup = async (e) => {
    e.preventDefault();
    if (!authEmail.trim() || !authPass.trim() || !authConfirm.trim()) {
      toast.error("Missing fields", "All fields are required.");
      return;
    }
    if (authPass !== authConfirm) {
      toast.error("Password mismatch", "Passwords do not match.");
      return;
    }
    if (authPass.length < 6) {
      toast.warning("Weak password", "Use at least 6 characters.");
      return;
    }
    setAuthLoading(true);
    try {
      const res = await axios.post("/api/auth/signup", {
        email: authEmail.trim(),
        password: authPass.trim()
      });
      if (res.data.success) {
        setAuthToken(res.data.token);
        setCurrentUser(res.data.user);
        setIsAuth(true);
        setShowLanding(false);
        toast.success("Account created!", "Welcome to SwiftPost.");
      }
    } catch (err) {
      toast.error("Registration failed", err.response?.data?.error || "Account creation failed.");
    } finally {
      setAuthLoading(false);
    }
  };

  function handleLogout() {
    setAuthToken("");
    setIsAuth(false);
    setShowLanding(true);
    setCurrentUser(null);
    setContacts([]);
    setFileHeaders([]);
    setSubject("");
    setBody("");
    setTemplates([]);
    setLogs([]);
    toast.info("Signed out", "Session ended successfully.");
  }

  // ── Insert HTML helper
  const insertHTML = (before, after = "") => {
    if (!bodyRef.current) return;
    const ta = bodyRef.current;
    const s = ta.selectionStart, end = ta.selectionEnd;
    const sel = ta.value.substring(s, end);
    setBody(ta.value.substring(0, s) + before + sel + after + ta.value.substring(end));
    setTimeout(() => { ta.focus(); ta.setSelectionRange(s + before.length, s + before.length + sel.length); }, 0);
  };

  const insertPlaceholder = h => {
    const token = `{${h}}`;
    if (activeField === "subject") setSubject(p => p + token);
    else insertHTML(token);
  };

  // ── File upload
  const handleFileUpload = e => {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: "binary" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (rows.length === 0) {
          toast.error("File is empty", "Selected sheet has no rows.");
          return;
        }

        let headers = [];
        let startRowIndex = 0;

        const firstRow = rows[0];
        const isEmailFormat = (str) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(str).trim());
        const hasEmailInFirstRow = firstRow.some((cell) => isEmailFormat(cell));

        if (hasEmailInFirstRow) {
          headers = firstRow.map((_, idx) => (idx === 0 ? "Email" : `Field_${idx + 1}`));
          startRowIndex = 0;
        } else {
          headers = firstRow.map((cell, idx) => {
            const val = String(cell || "").trim();
            return val ? val : `Field_${idx + 1}`;
          });
          startRowIndex = 1;
        }

        let emailColIdx = headers.findIndex(
          (h) => h.toLowerCase().includes("email") || h.toLowerCase().includes("mail")
        );
        if (emailColIdx === -1) emailColIdx = 0;

        const parsedContacts = [];
        for (let r = startRowIndex; r < rows.length; r++) {
          const row = rows[r];
          if (!row || row.length === 0) continue;

          const emailVal = String(row[emailColIdx] || "").trim();
          if (!emailVal) continue;

          const placeholders = {};
          headers.forEach((header, colIdx) => {
            placeholders[header] = String(row[colIdx] || "").trim();
          });

          parsedContacts.push({
            email: emailVal,
            valid: isEmailFormat(emailVal),
            placeholders: placeholders
          });
        }

        const res = await axios.post("/api/contacts/batch", {
          contacts: parsedContacts,
          overwrite: true
        });

        if (res.data.success) {
          fetchContacts();
          toast.success("Contacts imported!", `${parsedContacts.length} recipients loaded and saved.`);
        }
      } catch (err) {
        toast.error("Upload failed", err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  const downloadSample = () => toast.info("Download started", "BulkMail_Template.xlsx is being prepared.");

  // ── Contacts CRUD
  const filteredContacts = contacts.filter(c => {
    if (cFilter === "valid" && !c.valid) return false;
    if (cFilter === "invalid" && c.valid) return false;
    if (!cSearch) return true;
    const q = cSearch.toLowerCase();
    return c.email.toLowerCase().includes(q) || Object.values(c.placeholders).some(v => String(v).toLowerCase().includes(q));
  });
  const paginated = filteredContacts.slice((cPage - 1) * PER_PAGE, cPage * PER_PAGE);
  const totalPages = Math.ceil(filteredContacts.length / PER_PAGE) || 1;

  const addContact = async (e) => {
    e.preventDefault();
    if (!newEmail.trim()) return;
    try {
      const res = await axios.post("/api/contacts", {
        email: newEmail.trim(),
        valid: isValidEmail(newEmail.trim()),
        placeholders: { Email: newEmail.trim(), ...newFields }
      });
      if (res.data.success) {
        fetchContacts();
        setNewEmail(""); setNewFields({}); setShowAddForm(false);
        toast.success("Contact added", newEmail.trim());
      }
    } catch (err) {
      toast.error("Add failed", err.response?.data?.error || err.message);
    }
  };

  const deleteContact = async (id) => {
    try {
      const res = await axios.delete(`/api/contacts/${id}`);
      if (res.data.success) {
        setContacts(p => p.filter(c => (c._id || c.id) !== id));
        toast.warning("Removed", "Contact deleted from list.");
      }
    } catch (err) {
      toast.error("Delete failed", err.response?.data?.error || err.message);
    }
  };

  const saveEdit = async (id) => {
    try {
      const res = await axios.put(`/api/contacts/${id}`, {
        email: editEmail.trim(),
        valid: isValidEmail(editEmail.trim()),
        placeholders: { ...editFields, Email: editEmail.trim() }
      });
      if (res.data.success) {
        const updated = { ...res.data.data, id: res.data.data._id || res.data.data.id };
        setContacts(p => p.map(c => (c._id || c.id) === id ? updated : c));
        setEditingId(null);
        toast.success("Saved", "Contact updated.");
      }
    } catch (err) {
      toast.error("Save failed", err.response?.data?.error || err.message);
    }
  };

  const clearContacts = async () => {
    try {
      const res = await axios.delete("/api/contacts");
      if (res.data.success) {
        setContacts([]); setFileHeaders([]);
        toast.warning("Cleared", "All contacts removed.");
      }
    } catch (err) {
      toast.error("Clear failed", err.response?.data?.error || err.message);
    }
  };

  // ── Live interpolation
  const interpolate = txt => {
    if (!txt || contacts.length === 0) return txt || "";
    const r = contacts.find(c => c.valid) || contacts[0];
    let t = txt;
    Object.keys(r.placeholders).forEach(k => {
      t = t.replace(new RegExp(`{${k}}`, "g"), r.placeholders[k] || "");
    });
    return t;
  };

  // ── SMTP
  const saveSmtp = async (e) => {
    e.preventDefault();
    try {
      const res = await axios.post("/api/smtp", smtp);
      if (res.data.success) {
        toast.success("SMTP saved", `Configuration saved for ${smtp.host}`);
      }
    } catch (err) {
      toast.error("Save failed", err.response?.data?.error || err.message);
    }
  };

  const testSmtp = async (e) => {
    e.preventDefault();
    if (!testEmail.trim()) { toast.error("No email", "Enter a test recipient address."); return; }
    setTestLoading(true); setTestResult(null);
    try {
      const res = await axios.post("/sendmail", {
        email: [{ email: testEmail.trim(), placeholders: { Email: testEmail.trim() } }],
        sub: "SwiftPost SMTP Test Connection",
        msg: "Your SwiftPost connection is successfully verified and working. Happy sending!",
        html: true,
        smtpConfig: smtp
      });
      if (res.data.success) {
        setTestResult({ success: true, msg: "Test email sent! Check your inbox." });
        toast.success("SMTP verified!", `Test email sent to ${testEmail}`);
      } else {
        setTestResult({ success: false, msg: res.data.error || "SMTP failed" });
        toast.error("SMTP failed", res.data.error || "SMTP failed");
      }
    } catch (err) {
      const errorMsg = err.response?.data?.error || err.message;
      setTestResult({ success: false, msg: errorMsg });
      toast.error("SMTP failed", errorMsg);
    } finally {
      setTestLoading(false);
    }
  };

  // ── Templates
  const saveTemplate = async () => {
    if (!templateName.trim()) { toast.error("Name required", "Enter a template name."); return; }
    try {
      const res = await axios.post("/api/templates", {
        name: templateName.trim(),
        subject,
        body,
        isHtml: htmlMode
      });
      if (res.data.success) {
        setTemplates(p => [res.data.data, ...p]);
        setTemplateName(""); setShowSaveTplModal(false);
        toast.success("Template saved!", templateName.trim());
      }
    } catch (err) {
      toast.error("Save failed", err.response?.data?.error || err.message);
    }
  };

  const applyTemplate = tpl => {
    setSubject(tpl.subject); setBody(tpl.body); setHtmlMode(tpl.isHtml);
    setActiveTab("composer");
    toast.info("Template applied", tpl.name);
  };

  const deleteTemplate = async (id) => {
    try {
      const res = await axios.delete(`/api/templates/${id}`);
      if (res.data.success) {
        setTemplates(p => p.filter(t => t._id !== id));
        toast.warning("Deleted", "Template removed.");
      }
    } catch (err) {
      toast.error("Delete failed", err.response?.data?.error || err.message);
    }
  };

  // ── Campaign engine
  const startCampaign = async () => {
    const list = contacts.filter(c => c.valid);
    if (list.length === 0) { toast.error("No recipients", "Add valid email contacts first."); return; }
    if (!subject.trim()) { toast.error("Missing subject", "Enter a subject line."); return; }
    if (!body.trim()) { toast.error("Missing body", "Write the email content."); return; }

    setSending({ active: true, progress: 0, total: list.length, success: 0, fail: 0, paused: false });
    setConsoleLogs([]);
    campaignRef.current = true; pauseRef.current = false; idxRef.current = 0;

    logMsg(`Starting campaign — ${list.length} recipients`);
    logMsg(`SMTP: ${smtp.host || "unconfigured"}`);
    logMsg(`Throttle: ${throttle}ms`);

    let success = 0, fail = 0;
    const results = [];

    while (idxRef.current < list.length && campaignRef.current) {
      if (pauseRef.current) {
        logMsg("⏸ Paused — waiting...");
        setSending(p => ({ ...p, paused: true }));
        await new Promise(res => { const t = setInterval(() => { if (!pauseRef.current || !campaignRef.current) { clearInterval(t); res(); } }, 300); });
        if (!campaignRef.current) break;
        logMsg("▶ Resuming...");
        setSending(p => ({ ...p, paused: false }));
      }

      const c = list[idxRef.current];
      logMsg(`[${idxRef.current + 1}/${list.length}] → ${c.email}`);

      try {
        const res = await axios.post("/sendmail", {
          email: [{ email: c.email, placeholders: c.placeholders || {} }],
          sub: subject,
          msg: body,
          html: htmlMode
        });
        if (res.data.success && res.data.results?.[0]?.success) {
          success++;
          results.push({ email: c.email, success: true });
          logMsg(`✓ Delivered: ${c.email}`);
        } else {
          fail++;
          const errReason = res.data.results?.[0]?.error || res.data.error || "Rejected";
          results.push({ email: c.email, success: false, error: errReason });
          logMsg(`✗ Failed: ${c.email} — ${errReason}`);
        }
      } catch (err) {
        fail++;
        const errReason = err.response?.data?.error || err.message;
        results.push({ email: c.email, success: false, error: errReason });
        logMsg(`✗ Failed: ${c.email} — ${errReason}`);
      }

      idxRef.current++;
      // eslint-disable-next-line no-loop-func
      setSending(p => ({ ...p, progress: idxRef.current, success, fail }));

      if (idxRef.current < list.length && campaignRef.current)
        await new Promise(r => setTimeout(r, throttle));
    }

    const aborted = !campaignRef.current && idxRef.current < list.length;
    logMsg(aborted ? "🛑 Campaign cancelled." : "✅ Campaign complete!");

    const log = {
      date: new Date().toLocaleString(),
      subject,
      total: list.length,
      success,
      failure: fail,
      status: aborted ? "Cancelled" : "Completed",
      details: results
    };
    try {
      const res = await axios.post("/api/logs", log);
      if (res.data.success) {
        setLogs(p => [res.data.data, ...p]);
      }
    } catch (err) {
      console.error("Failed to save campaign log:", err.message);
    }
    setSending(p => ({ ...p, active: false }));
    campaignRef.current = false;

    if (!aborted) toast.success("Campaign done!", `${success} delivered, ${fail} failed.`);
  };

  const togglePause = () => {
    pauseRef.current = !pauseRef.current;
    setSending(p => ({ ...p, paused: pauseRef.current }));
  };

  const stopCampaign = () => { campaignRef.current = false; toast.warning("Campaign stopped", "Dispatch aborted by user."); };

  const clearLogs = async () => {
    try {
      const res = await axios.delete("/api/logs");
      if (res.data.success) {
        setLogs([]);
        toast.info("Logs cleared", "History has been removed.");
      }
    } catch (err) {
      toast.error("Clear failed", err.response?.data?.error || err.message);
    }
  };

  const askConfirm = (msg, cb) => setConfirm({ msg, cb });

  // ─────────────────────────────────────────────────────────────────────────────
  // LANDING PAGE
  // ─────────────────────────────────────────────────────────────────────────────
  if (!isAuth) {
    if (showLanding) return (
      <>
        <FontLoader />
        <GlobalStyles />
        <AlertSystem />
        <div className="sp-noise" style={{ cursor: "none" }}>
          <div ref={cursorRef} className="sp-cursor" />
          <div ref={ringRef} className="sp-cursor-ring" />

          {/* NAV */}
          <nav className="sp-nav">
            <span className="sp-logo" onClick={() => setShowLanding(true)}>Swift<span>Post</span></span>
            <ul className="sp-nav-links">
              <li><a href="#features">Features</a></li>
              <li><a href="#how">How it works</a></li>
              <li><a href="#pricing">Pricing</a></li>
              <li><button onClick={() => { setShowLanding(false); setAuthTab("login"); }}>Sign In</button></li>
              <li><button className="sp-nav-cta" onClick={() => { setShowLanding(false); setAuthTab("signup"); }}>Start Free</button></li>
            </ul>
          </nav>

          {/* HERO */}
          <section className="sp-hero">
            <div className="sp-hero-left">
              <p className="sp-eyebrow">Bulk Email Platform</p>
              <h1 className="sp-hero-title">
                <span className="sp-title-line"><span className={`sp-title-inner ${revealed ? "revealed" : ""}`}>Send</span></span>
                <span className="sp-title-line"><span className={`sp-title-inner ${revealed ? "revealed" : ""}`}>Millions.</span></span>
                <span className="sp-title-line"><span className={`sp-title-inner ${revealed ? "revealed" : ""}`}>Effortlessly.</span></span>
              </h1>
              <p className="sp-hero-desc">SwiftPost is the precision bulk mail engine built for modern teams — from <strong>transactional blasts</strong> to full campaign orchestration, at any scale.</p>
              <div className="sp-hero-actions">
                <button className="sp-btn-primary" onClick={() => { setShowLanding(false); setAuthTab("signup"); }}>Get Started Free</button>
                <a href="#how" className="sp-btn-ghost">See How It Works</a>
              </div>
              <div className="sp-stats">
                <div><div className="sp-stat-num">4.2B</div><div className="sp-stat-label">Emails Sent</div></div>
                <div className="sp-stat-div" />
                <div><div className="sp-stat-num">99.7%</div><div className="sp-stat-label">Deliverability</div></div>
                <div className="sp-stat-div" />
                <div><div className="sp-stat-num">18ms</div><div className="sp-stat-label">Avg Latency</div></div>
              </div>
            </div>
            <div className="sp-hero-right">
              <div className="sp-grid-bg" />
              <div className="sp-envelope-scene">
                <div className="sp-env-card"><div className="env-to">To: 84,200 recipients</div><div className="env-subject">Your Summer Sale Is Live 🎉</div><div className="env-meta">Campaign · Sent 2 min ago</div><div className="env-bar" style={{ width: "80%" }} /></div>
                <div className="sp-env-card"><div className="env-to">Transactional · 12,401/12,401</div><div className="env-subject">Order Confirmation</div><div className="env-meta">100% delivered · 0 bounced</div><div className="env-bar" style={{ width: "100%", background: "var(--sage)" }} /></div>
                <div className="sp-env-card"><div className="env-to">Newsletter · Power Users</div><div className="env-subject">This Month in Product</div><div className="env-meta">Open rate 38% · Queued</div><div className="env-bar" style={{ width: "38%", background: "var(--gold)" }} /></div>
                <div className="sp-env-card"><div className="env-to">Drip #3 of 5 · Automation</div><div className="env-subject">We thought you'd like this</div><div className="env-meta">Scheduled · 8:00 AM tomorrow</div><div className="env-bar" style={{ width: "60%" }} /></div>
              </div>
              <div className="sp-stamp">DELIVERED</div>
            </div>
          </section>

          {/* MARQUEE */}
          <div className="sp-marquee" aria-hidden="true">
            <div className="sp-marquee-track">
              {["BULK SEND","SMART SEGMENTATION","REAL-TIME ANALYTICS","DRIP CAMPAIGNS","A/B TESTING","SPF / DKIM / DMARC","WEBHOOKS","BULK SEND","SMART SEGMENTATION","REAL-TIME ANALYTICS","DRIP CAMPAIGNS","A/B TESTING","SPF / DKIM / DMARC","WEBHOOKS"].map((t,i) => (
                <span key={i}>{t}</span>
              )).reduce((a, e, i) => [...a, e, <span key={`d${i}`} className="dot">✦</span>], [])}
            </div>
          </div>

          {/* FEATURES */}
          <section className="sp-features" id="features">
            <p className="sp-section-label">Why SwiftPost</p>
            <h2 className="sp-section-title">Built for <em>scale,</em> refined for precision.</h2>
            <div className="sp-features-grid">
              {[
                ["⚡","01","Blazing Throughput","Send up to 10 million emails per hour across distributed nodes with auto-scaling infrastructure that never misses a beat."],
                ["🎯","02","Smart Segmentation","Build dynamic audiences from any data point — behavior, geography, purchase history, or custom attributes."],
                ["📊","03","Live Analytics","Opens, clicks, bounces, and unsubscribes — tracked in real time. Drill into any segment as sends flow."],
                ["🔐","04","Deliverability First","Automatic SPF, DKIM, and DMARC alignment. Reputation monitoring and suppression lists baked in."],
                ["🤖","05","Automation Flows","Trigger drip sequences, re-engagement campaigns, and event-driven sends via no-code builder or REST API."],
                ["🧪","06","A/B Testing","Test subject lines, send times, and content variants simultaneously. Auto-send the winner at any threshold."],
              ].map(([icon, num, title, desc]) => (
                <div key={num} className="sp-feat-card">
                  <span className="sp-feat-icon">{icon}</span>
                  <span className="sp-feat-num">{num}</span>
                  <h3 className="sp-feat-title">{title}</h3>
                  <p className="sp-feat-desc">{desc}</p>
                </div>
              ))}
            </div>
          </section>

          {/* HOW */}
          <section className="sp-how" id="how">
            <p className="sp-section-label" style={{ color: "var(--gold)" }}>How It Works</p>
            <h2 className="sp-section-title" style={{ color: "var(--cream)" }}>From list to inbox<br />in four steps.</h2>
            <div className="sp-steps">
              {[
                ["1","Import Your List","Upload CSV / XLSX. SwiftPost validates and deduplicates on ingestion."],
                ["2","Design Your Email","Use our HTML editor, write raw copy, or pull from the template library."],
                ["3","Set Targeting & Throttle","Segment your audience, control send pacing, and verify rendering live."],
                ["4","Track & Iterate","Watch sends land in real time. Review analytics and export reports."],
              ].map(([n, t, d]) => (
                <div key={n} className="sp-step">
                  <div className="sp-step-num">{n}</div>
                  <h3 className="sp-step-title">{t}</h3>
                  <p className="sp-step-desc">{d}</p>
                </div>
              ))}
            </div>
          </section>

          {/* PRICING */}
          <section className="sp-pricing" id="pricing">
            <p className="sp-section-label">Pricing</p>
            <h2 className="sp-section-title">Transparent rates.<br />No surprises.</h2>
            <div className="sp-pricing-grid">
              {[
                { label: "Starter", price: "$0", period: "Free Forever", features: ["50,000 emails/month","3 sending domains","Basic analytics","Email support"], featured: false },
                { label: "Growth",  price: "$0", period: "Free Forever", features: ["500,000 emails/month","Unlimited domains","Live analytics + A/B","Automation flows","Priority support"], featured: true },
                { label: "Enterprise", price: "Free", period: "Self-Hosted / SLA", features: ["Unlimited sends","Dedicated IPs","SLA guarantee","Dedicated support"], featured: false },
              ].map(({ label, price, period, features, featured }) => (
                <div key={label} className={`sp-price-card ${featured ? "featured" : ""}`}>
                  {featured && <div className="sp-price-badge">Most Popular</div>}
                  <p className="sp-price-plan">{label}</p>
                  <div className="sp-price-amount">{price}</div>
                  <p className="sp-price-period">{period}</p>
                  <ul className="sp-price-features">{features.map(f => <li key={f}>{f}</li>)}</ul>
                  <button className="sp-btn-plan" onClick={() => { setShowLanding(false); setAuthTab("signup"); }}>
                    {featured ? "Start Free Trial" : label === "Enterprise" ? "Contact Sales" : "Get Started"}
                  </button>
                </div>
              ))}
            </div>
          </section>

          {/* TESTIMONIALS */}
          <section className="sp-testimonials">
            <p className="sp-section-label">From Our Customers</p>
            <h2 className="sp-section-title">Trusted by teams<br />who send at scale.</h2>
            <div className="sp-testi-grid">
              {[
                ['"We migrated from another provider and cut our bounce rate in half overnight. The deliverability tooling is genuinely best-in-class."',"Priya Menon","Head of CRM · Shopwave"],
                ['"SwiftPost\'s API is the cleanest I\'ve worked with. We integrated our entire transactional flow in under a day with zero friction."',"Lucas Hartmann","Backend Lead · Finly"],
                ['"Real-time analytics changed how we run campaigns. Watching sends land live and reacting immediately is something our old tool couldn\'t do."',"Aisha Nkosi","Growth Manager · Nomad Studio"],
              ].map(([q, a, r]) => (
                <div key={a} className="sp-testi-card">
                  <p className="sp-testi-quote">{q}</p>
                  <div className="sp-testi-author">{a}</div>
                  <div className="sp-testi-role">{r}</div>
                </div>
              ))}
            </div>
          </section>

          {/* CTA */}
          <section className="sp-cta">
            <h2 className="sp-cta-title">Ready to <span>send</span><br />at scale?</h2>
            <p className="sp-cta-sub">Start for free. No credit card required. Your first 10,000 sends are on us.</p>
            <div className="sp-email-wrap">
              <input className="sp-email-input" type="email" placeholder="you@company.com" value={ctaEmail} onChange={e => setCtaEmail(e.target.value)} />
              <button className="sp-email-btn" onClick={() => { setAuthEmail(ctaEmail); setShowLanding(false); setAuthTab("signup"); }}>Get Early Access</button>
            </div>
          </section>

          {/* FOOTER */}
          <footer className="sp-footer">
            <div className="sp-footer-logo">Swift<span>Post</span></div>
            <ul className="sp-footer-links">
              <li><a href="#features">Docs</a></li>
              <li><a href="#features">Blog</a></li>
              <li><a href="#features">Status</a></li>
              <li><a href="#features">Privacy</a></li>
            </ul>
            <p className="sp-footer-copy">© 2026 SwiftPost Inc. All rights reserved.</p>
          </footer>
        </div>
      </>
    );

    // ── AUTH PAGE
    return (
      <>
        <FontLoader />
        <GlobalStyles />
        <AlertSystem />
        <div className="sp-auth sp-noise" style={{ cursor: "none" }}>
          <div ref={cursorRef} className="sp-cursor" />
          <div ref={ringRef} className="sp-cursor-ring" />

          <button onClick={() => setShowLanding(true)} style={{ position: "absolute", top: "1.5rem", left: "1.5rem", display: "flex", alignItems: "center", gap: ".4rem", background: "none", border: "none", cursor: "pointer", fontSize: ".75rem", fontWeight: 600, letterSpacing: ".12em", textTransform: "uppercase", color: "var(--muted)" }}>
            <ChevronLeft size={14} /> Back to Home
          </button>

          <div className="sp-auth-card">
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", marginBottom: "1.75rem", gap: ".75rem" }}>
              <div style={{ width: 48, height: 48, borderRadius: 4, background: "var(--rust)", display: "flex", alignItems: "center", justifyContent: "center", color: "var(--cream)" }}>
                <Mail size={22} />
              </div>
              <div className="sp-auth-title">Swift<span>Post</span></div>
              <div className="sp-auth-subtitle">Database Authenticated Engine</div>
            </div>

            <div style={{ display: "flex", borderBottom: "1px solid var(--border)", marginBottom: "1.5rem" }}>
              <button className={`sp-auth-tab ${authTab === "login" ? "active" : ""}`} onClick={() => setAuthTab("login")}>Sign In</button>
              <button className={`sp-auth-tab ${authTab === "signup" ? "active" : ""}`} onClick={() => setAuthTab("signup")}>Register</button>
            </div>

            {authTab === "login" ? (
              <form onSubmit={handleLogin} style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
                <div>
                  <label className="sp-auth-label">Email Address</label>
                  <div style={{ position: "relative" }}>
                    <Mail size={14} style={{ position: "absolute", left: ".9rem", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                    <input className="sp-auth-input" type="email" required value={authEmail} onChange={e => setAuthEmail(e.target.value)} placeholder="you@company.com" />
                  </div>
                </div>
                <div>
                  <label className="sp-auth-label">Password</label>
                  <div style={{ position: "relative" }}>
                    <Lock size={14} style={{ position: "absolute", left: ".9rem", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                    <input className="sp-auth-input" type={showAuthPass ? "text" : "password"} required value={authPass} onChange={e => setAuthPass(e.target.value)} placeholder="••••••••" />
                    <button type="button" onClick={() => setShowAuthPass(p => !p)} style={{ position: "absolute", right: ".9rem", top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", color: "var(--muted)" }}><Eye size={14} /></button>
                  </div>
                </div>
                <button type="submit" disabled={authLoading} className="sp-auth-submit">
                  {authLoading ? <><RefreshCw size={14} className="sp-spin" /> Authenticating…</> : "Authenticate Securely"}
                </button>
              </form>
            ) : (
              <form onSubmit={handleSignup} style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
                <div>
                  <label className="sp-auth-label">Email Address</label>
                  <div style={{ position: "relative" }}>
                    <Mail size={14} style={{ position: "absolute", left: ".9rem", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                    <input className="sp-auth-input" type="email" required value={authEmail} onChange={e => setAuthEmail(e.target.value)} placeholder="you@company.com" />
                  </div>
                </div>
                <div>
                  <label className="sp-auth-label">Password</label>
                  <div style={{ position: "relative" }}>
                    <Lock size={14} style={{ position: "absolute", left: ".9rem", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                    <input className="sp-auth-input" type={showAuthPass ? "text" : "password"} required value={authPass} onChange={e => setAuthPass(e.target.value)} placeholder="Create a password…" />
                    <button type="button" onClick={() => setShowAuthPass(p => !p)} style={{ position: "absolute", right: ".9rem", top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", color: "var(--muted)" }}><Eye size={14} /></button>
                  </div>
                </div>
                <div>
                  <label className="sp-auth-label">Confirm Password</label>
                  <div style={{ position: "relative" }}>
                    <Lock size={14} style={{ position: "absolute", left: ".9rem", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                    <input className="sp-auth-input" type="password" required value={authConfirm} onChange={e => setAuthConfirm(e.target.value)} placeholder="Verify password…" />
                  </div>
                </div>
                <button type="submit" disabled={authLoading} className="sp-auth-submit">
                  {authLoading ? <><RefreshCw size={14} className="sp-spin" /> Creating account…</> : "Create Account"}
                </button>
              </form>
            )}

            <div style={{ marginTop: "1.5rem", textAlign: "center", fontSize: ".7rem", color: "var(--muted)", letterSpacing: ".1em", textTransform: "uppercase" }}>
              Cluster: <span style={{ color: "#4ade80", fontWeight: 700 }}>Online</span>
            </div>
          </div>
        </div>
      </>
    );
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // DASHBOARD
  // ─────────────────────────────────────────────────────────────────────────────
  const NAV_ITEMS = [
    { key: "composer",  label: "Composer",         icon: <Mail size={15} /> },
    { key: "contacts",  label: "Recipients",        icon: <Users size={15} />, badge: contacts.length || null },
    { key: "templates", label: "Templates",         icon: <FileText size={15} />, badge: templates.length || null },
    { key: "smtp",      label: "SMTP Server",       icon: <Sliders size={15} /> },
    { key: "logs",      label: "Campaign History",  icon: <History size={15} />, badge: logs.length || null },
  ];

  const TAB_TITLE = { composer: "Campaign Composer", contacts: "Recipients Manager", templates: "Email Templates", smtp: "SMTP Configuration", logs: "Dispatch History" };

  return (
    <>
      <FontLoader />
      <GlobalStyles />
      <AlertSystem />
      {confirm && <ConfirmDialog msg={confirm.msg} onConfirm={() => { confirm.cb(); setConfirm(null); }} onCancel={() => setConfirm(null)} />}

      <div className="sp-dash">
        {/* ── SIDEBAR ── */}
        <aside className="sp-sidebar">
          <div className="sp-sidebar-logo">
            <div className="logo-mark">Swift<span>Post</span></div>
            <div className="logo-tag">System Panel</div>
          </div>

          <nav className="sp-sidebar-nav">
            {NAV_ITEMS.map(({ key, label, icon, badge }) => (
              <button key={key} className={`sp-nav-item ${activeTab === key ? "active" : ""}`} onClick={() => setActiveTab(key)}>
                <div className="nav-left">{icon} {label}</div>
                {badge ? <span className="sp-badge">{badge}</span> : null}
              </button>
            ))}
          </nav>

          <div className="sp-sidebar-footer">
            {currentUser && (
              <div style={{ display: "flex", alignItems: "center", gap: ".6rem", padding: ".75rem", background: "rgba(245,240,232,.03)", border: "1px solid rgba(245,240,232,.06)", borderRadius: 2, marginBottom: ".75rem" }}>
                <UserCheck size={14} style={{ color: "var(--rust)", flexShrink: 0 }} />
                <div style={{ flex: 1, overflow: "hidden" }}>
                  <div style={{ fontSize: ".75rem", fontWeight: 600, color: "var(--cream)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{currentUser.email}</div>
                  <div style={{ fontSize: ".6rem", color: "rgba(245,240,232,.3)", textTransform: "uppercase", letterSpacing: ".15em", marginTop: 2 }}>Session Active</div>
                </div>
                <button onClick={handleLogout} title="Sign Out" style={{ background: "none", border: "none", cursor: "pointer", color: "rgba(245,240,232,.3)", padding: 4 }}>
                  <LogOut size={13} />
                </button>
              </div>
            )}

            <div style={{ display: "flex", alignItems: "center", gap: ".6rem", padding: ".6rem .75rem", background: "rgba(245,240,232,.02)", borderRadius: 2, border: "1px solid rgba(245,240,232,.05)" }}>
              <div className={`sp-status-dot ${serverStatus}`} />
              <span style={{ fontSize: ".72rem", fontWeight: 600, color: "rgba(245,240,232,.4)", letterSpacing: ".05em" }}>
                {serverStatus === "online" ? "Server Online" : serverStatus === "offline" ? "Server Offline" : "Checking…"}
              </span>
              <button onClick={() => { setServerStatus("checking"); setTimeout(() => setServerStatus("online"), 600); }} style={{ marginLeft: "auto", background: "none", border: "none", cursor: "pointer", color: "rgba(245,240,232,.25)" }}>
                <RefreshCw size={11} />
              </button>
            </div>
          </div>
        </aside>

        {/* ── MAIN ── */}
        <main className="sp-main">
          <header className="sp-topbar">
            <div>
              <span style={{ fontSize: ".65rem", fontWeight: 600, letterSpacing: ".2em", textTransform: "uppercase", color: "rgba(245,240,232,.25)" }}>Workspace / </span>
              <span className="sp-topbar-title">{TAB_TITLE[activeTab]}</span>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: "1rem" }}>
              <div style={{ display: "flex", alignItems: "center", gap: ".5rem", padding: ".35rem .85rem", background: "rgba(245,240,232,.04)", border: "1px solid rgba(245,240,232,.06)", borderRadius: 2, fontSize: ".72rem", color: "rgba(245,240,232,.4)" }}>
                <Database size={12} style={{ color: "var(--rust)" }} />
                SMTP: <strong style={{ color: "rgba(245,240,232,.7)" }}>{smtp.host || "Not configured"}</strong>
              </div>
              <span style={{ fontSize: ".72rem", color: "rgba(245,240,232,.25)" }}>{new Date().toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric" })}</span>
            </div>
          </header>

          <div className="sp-content">

            {/* ══ TAB: COMPOSER ══════════════════════════════════════════════ */}
            {activeTab === "composer" && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 420px", gap: "1.5rem" }} className="sp-fade-in">
                {/* LEFT */}
                <div style={{ display: "flex", flexDirection: "column", gap: "1.5rem" }}>
                  <div className="sp-card">
                    <div className="sp-card-title"><Mail size={16} style={{ color: "var(--rust)" }} /> Compose Message</div>

                    {contacts.length === 0 && (
                      <div style={{ display: "flex", gap: ".75rem", padding: "1rem", background: "rgba(251,191,36,.05)", border: "1px solid rgba(251,191,36,.15)", borderRadius: 2, marginBottom: "1.25rem" }}>
                        <AlertCircle size={14} style={{ color: "#fbbf24", flexShrink: 0, marginTop: 1 }} />
                        <div style={{ fontSize: ".78rem" }}>
                          <strong style={{ color: "#fbbf24" }}>No contacts imported.</strong>
                          <span style={{ color: "rgba(245,240,232,.4)", marginLeft: ".35rem" }}>Add recipients in the Recipients tab before launching.</span>
                        </div>
                      </div>
                    )}

                    <div style={{ display: "flex", flexDirection: "column", gap: "1.1rem" }}>
                      <div>
                        <label className="sp-label">Subject Line</label>
                        <input className="sp-input" value={subject} onChange={e => setSubject(e.target.value)} onFocus={() => setActiveField("subject")} placeholder="Enter campaign subject…" />
                      </div>

                      {fileHeaders.length > 0 && (
                        <div>
                          <label className="sp-label">Insert Merge Tags</label>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: ".4rem" }}>
                            {fileHeaders.map(h => (
                              <button key={h} className="sp-placeholder" onClick={() => insertPlaceholder(h)}>+ &#123;{h}&#125;</button>
                            ))}
                          </div>
                        </div>
                      )}

                      <div>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: ".4rem" }}>
                          <label className="sp-label" style={{ marginBottom: 0 }}>Email Body</label>
                          <button onClick={() => setHtmlMode(p => !p)} style={{ padding: "2px 10px", borderRadius: 2, border: `1px solid ${htmlMode ? "rgba(74,222,128,.2)" : "rgba(245,240,232,.08)"}`, background: htmlMode ? "rgba(74,222,128,.08)" : "transparent", color: htmlMode ? "#4ade80" : "rgba(245,240,232,.35)", fontSize: ".65rem", fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", cursor: "pointer" }}>
                            {htmlMode ? "HTML" : "Plain Text"}
                          </button>
                        </div>

                        {htmlMode && (
                          <div className="sp-toolbar">
                            {[["B","<strong>","</strong>"],["I","<em>","</em>"],["H1","<h1>","</h1>"],["H2","<h2>","</h2>"],["Link","<a href=\"\">","</a>"],["BR","<br/>",""]].map(([l, b, a]) => (
                              <button key={l} className="sp-toolbar-btn" onClick={() => insertHTML(b, a)}>{l}</button>
                            ))}
                          </div>
                        )}

                        <textarea ref={bodyRef} value={body} onChange={e => setBody(e.target.value)} onFocus={() => setActiveField("body")} placeholder={htmlMode ? "<p>Hello {Name}, ...</p>" : "Enter plain text…"} className={`sp-input sp-textarea ${htmlMode ? "" : ""}`} style={{ minHeight: 280, display: "block", width: "100%", borderRadius: htmlMode ? "0 0 2px 2px" : 2, fontFamily: htmlMode ? "'JetBrains Mono', 'Fira Code', monospace" : undefined }} />
                      </div>

                      <div style={{ paddingTop: "1rem", borderTop: "1px solid rgba(245,240,232,.06)", display: "flex", alignItems: "center", justifyContent: "space-between", gap: "1rem" }}>
                        <div>
                          <label className="sp-label">Throttle Delay — <span style={{ color: "var(--rust)", fontWeight: 700 }}>{(throttle/1000).toFixed(1)}s</span></label>
                          <input type="range" min={100} max={10000} step={100} value={throttle} onChange={e => setThrottle(+e.target.value)} style={{ width: 140 }} />
                        </div>
                        <div style={{ display: "flex", gap: ".6rem" }}>
                          <button className="sp-btn-dash sp-btn-ink" onClick={() => setShowSaveTplModal(true)} disabled={!subject || !body}>
                            <FileText size={13} /> Save Draft
                          </button>
                          <button className={`sp-btn-dash sp-btn-rust ${(!contacts.filter(c=>c.valid).length || !subject || !body) ? "sp-btn-disabled" : ""}`} onClick={startCampaign} disabled={!contacts.filter(c=>c.valid).length || !subject || !body}>
                            <Send size={13} /> Send Campaign
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                {/* RIGHT — PREVIEW */}
                <div className="sp-card" style={{ display: "flex", flexDirection: "column", gap: ".75rem" }}>
                  <div className="sp-card-title"><Eye size={16} style={{ color: "var(--rust)" }} /> Live Preview</div>
                  <p style={{ fontSize: ".75rem", color: "rgba(245,240,232,.3)", lineHeight: 1.5 }}>Resolved for the first valid recipient.</p>
                  <div className="sp-browser" style={{ flex: 1 }}>
                    <div className="sp-browser-bar">
                      <div className="sp-browser-dots">
                        <div className="sp-browser-dot" style={{ background: "#f87171" }} />
                        <div className="sp-browser-dot" style={{ background: "#fbbf24" }} />
                        <div className="sp-browser-dot" style={{ background: "#4ade80" }} />
                      </div>
                      <span style={{ fontSize: ".65rem", color: "rgba(245,240,232,.2)", letterSpacing: ".1em", textTransform: "uppercase", marginLeft: "auto" }}>Mail Simulator</span>
                    </div>
                    <div style={{ background: "#0d1117", padding: ".75rem 1rem", borderBottom: "1px solid rgba(245,240,232,.05)" }}>
                      {[
                        ["From", smtp.fromEmail || smtp.user || "sender@example.com"],
                        ["To",   contacts.length ? (contacts.find(c=>c.valid)||contacts[0]).email : "recipient@example.com"],
                        ["Sub",  subject ? interpolate(subject) : "(No Subject)"],
                      ].map(([k, v]) => (
                        <div key={k} style={{ display: "flex", gap: ".6rem", marginBottom: ".3rem", fontSize: ".73rem" }}>
                          <span style={{ color: "rgba(245,240,232,.25)", fontWeight: 700, width: 32, flexShrink: 0 }}>{k}:</span>
                          <span style={{ color: k === "To" ? "var(--rust)" : "rgba(245,240,232,.7)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{v}</span>
                        </div>
                      ))}
                    </div>
                    <div className="sp-browser-body" style={{ minHeight: 240 }}>
                      {htmlMode
                        ? <div style={{ fontSize: 13, lineHeight: 1.6 }} dangerouslySetInnerHTML={{ __html: interpolate(body) || "<p style='color:#94a3b8;font-style:italic;'>Start writing to see a preview here…</p>" }} />
                        : <pre style={{ whiteSpace: "pre-wrap", fontFamily: "inherit", fontSize: 13, color: "#1e293b" }}>{interpolate(body) || "Your content appears here…"}</pre>
                      }
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* ══ TAB: CONTACTS ══════════════════════════════════════════════ */}
            {activeTab === "contacts" && (
              <div style={{ display: "flex", flexDirection: "column", gap: "1.5rem" }} className="sp-fade-in">
                {/* STATS */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: "1rem" }} className="sp-stagger">
                  {[
                    [contacts.length, "Total Recipients", <Users size={20} />, "var(--rust)"],
                    [contacts.filter(c=>c.valid).length, "Valid Emails", <CheckCircle2 size={20} />, "#4ade80"],
                    [contacts.filter(c=>!c.valid).length, "Invalid Emails", <XCircle size={20} />, "#f87171"],
                    [fileHeaders.length, "Campaign Variables", <FileSpreadsheet size={20} />, "var(--gold)"],
                  ].map(([n, l, icon, color]) => (
                    <div key={l} className="sp-stat-card">
                      <div>
                        <div className="sp-label">{l}</div>
                        <div className="sp-stat-big" style={{ color }}>{n}</div>
                      </div>
                      <div style={{ width: 44, height: 44, borderRadius: 2, background: `${color}15`, display: "flex", alignItems: "center", justifyContent: "center", color }}>{icon}</div>
                    </div>
                  ))}
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "300px 1fr", gap: "1.5rem" }}>
                  {/* UPLOAD */}
                  <div className="sp-card" style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
                    <div className="sp-card-title"><Upload size={15} style={{ color: "var(--rust)" }} /> File Import</div>
                    <p style={{ fontSize: ".78rem", color: "rgba(245,240,232,.35)", lineHeight: 1.6 }}>Upload XLSX or CSV. Standard mapping finds the Email column automatically.</p>
                    <div className="sp-dropzone">
                      <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} />
                      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: ".6rem" }}>
                        <Upload size={24} style={{ color: "rgba(245,240,232,.2)" }} />
                        <span style={{ fontSize: ".78rem", color: "rgba(245,240,232,.45)" }}>Choose XLSX / CSV</span>
                        <span style={{ fontSize: ".68rem", color: "rgba(245,240,232,.2)" }}>or drag & drop</span>
                      </div>
                    </div>
                    <button className="sp-btn-dash sp-btn-ink" onClick={downloadSample} style={{ width: "100%", justifyContent: "center" }}>
                      <Download size={13} style={{ color: "var(--rust)" }} /> Download Sample
                    </button>
                    {contacts.length > 0 && (
                      <button className="sp-btn-dash sp-btn-danger" onClick={() => askConfirm("Clear all contacts from the list?", clearContacts)} style={{ width: "100%", justifyContent: "center" }}>
                        <Trash2 size={13} /> Clear All
                      </button>
                    )}
                  </div>

                  {/* TABLE */}
                  <div className="sp-card">
                    <div style={{ display: "flex", alignItems: "center", gap: ".75rem", marginBottom: "1rem", flexWrap: "wrap" }}>
                      <div style={{ position: "relative", flex: 1, minWidth: 180 }}>
                        <Search size={13} style={{ position: "absolute", left: ".75rem", top: "50%", transform: "translateY(-50%)", color: "rgba(245,240,232,.2)" }} />
                        <input className="sp-input" style={{ paddingLeft: "2.2rem" }} placeholder="Search recipients…" value={cSearch} onChange={e => { setCSearch(e.target.value); setCPage(1); }} />
                      </div>
                      <div style={{ display: "flex", gap: 2, background: "rgba(245,240,232,.04)", borderRadius: 2, padding: 2, border: "1px solid rgba(245,240,232,.06)" }}>
                        {["all","valid","invalid"].map(f => (
                          <button key={f} onClick={() => { setCFilter(f); setCPage(1); }} style={{ padding: ".3rem .75rem", borderRadius: 2, border: "none", cursor: "pointer", fontSize: ".65rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: ".1em", background: cFilter === f ? (f==="valid" ? "#4ade8022" : f==="invalid" ? "#f8717122" : "rgba(201,75,42,.15)") : "transparent", color: cFilter === f ? (f==="valid" ? "#4ade80" : f==="invalid" ? "#f87171" : "var(--rust)") : "rgba(245,240,232,.3)", transition: "all .15s" }}>
                            {f}
                          </button>
                        ))}
                      </div>
                      <button className="sp-btn-dash sp-btn-rust" onClick={() => setShowAddForm(p => !p)}>
                        <Plus size={13} /> Add
                      </button>
                    </div>

                    {showAddForm && (
                      <form onSubmit={addContact} style={{ marginBottom: "1rem", padding: "1rem", background: "rgba(201,75,42,.05)", border: "1px solid rgba(201,75,42,.15)", borderRadius: 2 }}>
                        <div style={{ fontSize: ".65rem", fontWeight: 700, letterSpacing: ".15em", textTransform: "uppercase", color: "var(--rust)", marginBottom: ".75rem" }}>Add Recipient</div>
                        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: ".6rem", marginBottom: ".75rem" }}>
                          <div>
                            <label className="sp-label">Email *</label>
                            <input className="sp-input" type="email" required value={newEmail} onChange={e => setNewEmail(e.target.value)} placeholder="user@example.com" />
                          </div>
                          {fileHeaders.filter(h => h.toLowerCase() !== "email").map(h => (
                            <div key={h}>
                              <label className="sp-label">{h}</label>
                              <input className="sp-input" value={newFields[h]||""} onChange={e => setNewFields(p => ({ ...p, [h]: e.target.value }))} placeholder={h} />
                            </div>
                          ))}
                        </div>
                        <div style={{ display: "flex", gap: ".5rem", justifyContent: "flex-end" }}>
                          <button type="button" className="sp-btn-dash sp-btn-ink" onClick={() => setShowAddForm(false)}>Cancel</button>
                          <button type="submit" className="sp-btn-dash sp-btn-rust">Add to List</button>
                        </div>
                      </form>
                    )}

                    <div style={{ border: "1px solid rgba(245,240,232,.05)", borderRadius: 2, overflow: "hidden" }}>
                      <table className="sp-table">
                        <thead>
                          <tr>
                            <th>Status</th>
                            <th>Email</th>
                            {fileHeaders.filter(h => h.toLowerCase()!=="email").slice(0,3).map(h => <th key={h}>{h}</th>)}
                            <th style={{ textAlign: "right" }}>Actions</th>
                          </tr>
                        </thead>
                        <tbody>
                          {paginated.length === 0 ? (
                            <tr><td colSpan={10} style={{ textAlign: "center", padding: "2rem", color: "rgba(245,240,232,.2)", fontStyle: "italic" }}>No contacts found.</td></tr>
                          ) : paginated.map(c => (
                            <tr key={c.id}>
                              <td>
                                <span className={`sp-chip ${c.valid ? "sp-chip-green" : "sp-chip-red"}`}>
                                  <span style={{ width: 6, height: 6, borderRadius: "50%", background: "currentColor" }} />
                                  {c.valid ? "Valid" : "Invalid"}
                                </span>
                              </td>
                              <td style={{ fontWeight: 600 }}>
                                {editingId === c.id
                                  ? <input className="sp-input" style={{ padding: ".3rem .5rem", fontSize: ".8rem" }} value={editEmail} onChange={e => setEditEmail(e.target.value)} />
                                  : c.email}
                              </td>
                              {fileHeaders.filter(h=>h.toLowerCase()!=="email").slice(0,3).map(h => (
                                <td key={h} style={{ color: "rgba(245,240,232,.4)" }}>
                                  {editingId === c.id
                                    ? <input className="sp-input" style={{ padding: ".3rem .5rem", fontSize: ".8rem" }} value={editFields[h]||""} onChange={e => setEditFields(p => ({ ...p, [h]: e.target.value }))} />
                                    : c.placeholders[h]||"—"}
                                </td>
                              ))}
                              <td style={{ textAlign: "right" }}>
                                {editingId === c.id ? (
                                  <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                                    <button onClick={() => saveEdit(c.id)} style={{ width: 28, height: 28, borderRadius: 2, border: "1px solid rgba(74,222,128,.2)", background: "rgba(74,222,128,.08)", color: "#4ade80", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}><Check size={12} /></button>
                                    <button onClick={() => setEditingId(null)} style={{ width: 28, height: 28, borderRadius: 2, border: "1px solid rgba(245,240,232,.08)", background: "transparent", color: "rgba(245,240,232,.3)", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}><X size={12} /></button>
                                  </div>
                                ) : (
                                  <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                                    <button onClick={() => { setEditingId(c.id); setEditEmail(c.email); setEditFields({...c.placeholders}); }} style={{ width: 28, height: 28, borderRadius: 2, border: "1px solid rgba(201,75,42,.15)", background: "rgba(201,75,42,.06)", color: "var(--rust)", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}><Edit2 size={12} /></button>
                                    <button onClick={() => deleteContact(c.id)} style={{ width: 28, height: 28, borderRadius: 2, border: "1px solid rgba(248,113,113,.15)", background: "rgba(248,113,113,.06)", color: "#f87171", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}><Trash2 size={12} /></button>
                                  </div>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {totalPages > 1 && (
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: ".75rem" }}>
                        <span style={{ fontSize: ".68rem", color: "rgba(245,240,232,.25)", letterSpacing: ".1em", textTransform: "uppercase" }}>
                          {(cPage-1)*PER_PAGE+1}–{Math.min(cPage*PER_PAGE, filteredContacts.length)} of {filteredContacts.length}
                        </span>
                        <div style={{ display: "flex", gap: ".4rem", alignItems: "center" }}>
                          <button className="sp-page-btn" onClick={() => setCPage(p=>Math.max(1,p-1))} disabled={cPage===1}><ChevronLeft size={13} /></button>
                          <span style={{ fontSize: ".72rem", fontWeight: 700, color: "var(--rust)", padding: "0 .5rem" }}>{cPage}/{totalPages}</span>
                          <button className="sp-page-btn" onClick={() => setCPage(p=>Math.min(totalPages,p+1))} disabled={cPage===totalPages}><ChevronRight size={13} /></button>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}

            {/* ══ TAB: TEMPLATES ══════════════════════════════════════════════ */}
            {activeTab === "templates" && (
              <div className="sp-fade-in">
                <div style={{ marginBottom: "1.5rem" }}>
                  <h2 style={{ fontFamily: "'Clash Display', sans-serif", fontWeight: 500, fontSize: "1.5rem", letterSpacing: "-.02em", color: "var(--cream)" }}>Template Library</h2>
                  <p style={{ fontSize: ".78rem", color: "rgba(245,240,232,.3)", marginTop: ".3rem" }}>Save and apply email blueprints from the Composer.</p>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "1rem" }}>
                  {templates.map(tpl => (
                    <div key={tpl._id} className="sp-tpl-card sp-lift">
                      <div>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: ".6rem" }}>
                          <span style={{ fontFamily: "'Clash Display', sans-serif", fontWeight: 500, fontSize: ".95rem", color: "var(--cream)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: "75%" }}>{tpl.name}</span>
                          <span style={{ fontSize: ".6rem", fontWeight: 700, padding: "2px 8px", borderRadius: 2, textTransform: "uppercase", letterSpacing: ".1em", background: tpl.isHtml ? "rgba(74,222,128,.08)" : "rgba(245,240,232,.04)", color: tpl.isHtml ? "#4ade80" : "rgba(245,240,232,.3)", border: `1px solid ${tpl.isHtml ? "rgba(74,222,128,.2)" : "rgba(245,240,232,.06)"}` }}>{tpl.isHtml ? "HTML" : "TXT"}</span>
                        </div>
                        <p style={{ fontSize: ".72rem", color: "rgba(245,240,232,.3)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginBottom: ".5rem" }}>Subject: {tpl.subject}</p>
                        <p style={{ fontSize: ".72rem", color: "rgba(245,240,232,.2)", lineHeight: 1.5, display: "-webkit-box", WebkitLineClamp: 3, WebkitBoxOrient: "vertical", overflow: "hidden" }}>{tpl.body.replace(/<[^>]*>/g,"")}</p>
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderTop: "1px solid rgba(245,240,232,.05)", paddingTop: ".75rem" }}>
                        <button className="sp-btn-dash sp-btn-danger" onClick={() => deleteTemplate(tpl._id)}><Trash2 size={12} /></button>
                        <button className="sp-btn-dash sp-btn-rust" onClick={() => applyTemplate(tpl)}>Apply <ArrowRight size={12} /></button>
                      </div>
                    </div>
                  ))}
                  {templates.length === 0 && (
                    <div style={{ gridColumn: "1/-1", textAlign: "center", padding: "4rem", border: "2px dashed rgba(245,240,232,.06)", borderRadius: 4 }}>
                      <FileText size={36} style={{ color: "rgba(245,240,232,.1)", margin: "0 auto 1rem" }} />
                      <div style={{ fontFamily: "'Clash Display', sans-serif", fontSize: "1.1rem", color: "rgba(245,240,232,.3)" }}>No templates saved yet</div>
                      <div style={{ fontSize: ".78rem", color: "rgba(245,240,232,.2)", marginTop: ".4rem" }}>Save a draft from the Composer to see it here.</div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* ══ TAB: SMTP ══════════════════════════════════════════════ */}
            {activeTab === "smtp" && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 380px", gap: "1.5rem" }} className="sp-fade-in">
                <form className="sp-card" onSubmit={saveSmtp}>
                  <div className="sp-card-title"><Sliders size={16} style={{ color: "var(--rust)" }} /> SMTP Server Configuration</div>
                  <p style={{ fontSize: ".78rem", color: "rgba(245,240,232,.3)", lineHeight: 1.6, marginBottom: "1.5rem" }}>Configure your outgoing email node. Compatible with Gmail, Outlook, Amazon SES, SendGrid, and any custom SMTP server.</p>
                  <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 120px", gap: ".75rem" }}>
                      <div><label className="sp-label">SMTP Host</label><input className="sp-input" value={smtp.host} onChange={e => setSmtp(p=>({...p,host:e.target.value}))} placeholder="smtp.gmail.com" /></div>
                      <div><label className="sp-label">Port</label><input className="sp-input" type="number" value={smtp.port} onChange={e => setSmtp(p=>({...p,port:+e.target.value}))} placeholder="587" /></div>
                    </div>
                    <label style={{ display: "flex", alignItems: "center", gap: ".6rem", cursor: "pointer", padding: ".65rem .85rem", background: "rgba(245,240,232,.02)", border: "1px solid rgba(245,240,232,.06)", borderRadius: 2 }}>
                      <input type="checkbox" checked={smtp.secure} onChange={e=>setSmtp(p=>({...p,secure:e.target.checked}))} style={{ accentColor: "var(--rust)" }} />
                      <span style={{ fontSize: ".78rem", color: "rgba(245,240,232,.6)" }}>Use SSL/TLS (Port 465)</span>
                    </label>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: ".75rem" }}>
                      <div><label className="sp-label">Username / Email</label><input className="sp-input" value={smtp.user} onChange={e=>setSmtp(p=>({...p,user:e.target.value}))} placeholder="you@gmail.com" /></div>
                      <div>
                        <label className="sp-label">Password / App Token</label>
                        <div style={{ position: "relative" }}>
                          <input className="sp-input" type={showPass?"text":"password"} value={smtp.pass} onChange={e=>setSmtp(p=>({...p,pass:e.target.value}))} placeholder="••••••••••••" style={{ paddingRight: "2.5rem" }} />
                          <button type="button" onClick={()=>setShowPass(p=>!p)} style={{ position: "absolute", right: ".75rem", top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", color: "rgba(245,240,232,.2)" }}><Eye size={14} /></button>
                        </div>
                      </div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: ".75rem" }}>
                      <div><label className="sp-label">Sender Display Name</label><input className="sp-input" value={smtp.fromName} onChange={e=>setSmtp(p=>({...p,fromName:e.target.value}))} placeholder="SwiftPost Team" /></div>
                      <div><label className="sp-label">Sender Email</label><input className="sp-input" type="email" value={smtp.fromEmail} onChange={e=>setSmtp(p=>({...p,fromEmail:e.target.value}))} placeholder="noreply@company.com" /></div>
                    </div>
                    <div style={{ display: "flex", alignItems: "flex-start", gap: ".6rem", padding: ".75rem .85rem", background: "rgba(201,75,42,.05)", border: "1px solid rgba(201,75,42,.1)", borderRadius: 2 }}>
                      <Info size={13} style={{ color: "var(--rust)", flexShrink: 0, marginTop: 2 }} />
                      <p style={{ fontSize: ".73rem", color: "rgba(245,240,232,.35)", lineHeight: 1.6 }}>For Gmail, generate a 16-character <strong style={{ color: "rgba(245,240,232,.6)" }}>App Password</strong> from Google Account Security instead of your account password.</p>
                    </div>
                    <div style={{ display: "flex", justifyContent: "flex-end" }}>
                      <button type="submit" className="sp-btn-dash sp-btn-rust">
                        <Database size={13} /> Save Configuration
                      </button>
                    </div>
                  </div>
                </form>

                {/* TEST PANEL */}
                <div className="sp-card" style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
                  <div className="sp-card-title"><Shield size={16} style={{ color: "var(--rust)" }} /> Verify Connection</div>
                  <p style={{ fontSize: ".78rem", color: "rgba(245,240,232,.3)", lineHeight: 1.6 }}>Send a test message to verify your SMTP credentials and configuration.</p>
                  <form onSubmit={testSmtp} style={{ display: "flex", flexDirection: "column", gap: ".75rem" }}>
                    <div>
                      <label className="sp-label">Test Recipient</label>
                      <input className="sp-input" type="email" required value={testEmail} onChange={e=>setTestEmail(e.target.value)} placeholder="test@example.com" />
                    </div>
                    <button type="submit" className="sp-btn-dash sp-btn-rust" disabled={testLoading} style={{ justifyContent: "center" }}>
                      {testLoading ? <><RefreshCw size={13} className="sp-spin" /> Verifying…</> : <><Send size={13} /> Send Test Email</>}
                    </button>
                  </form>

                  {testResult && (
                    <div style={{ display: "flex", gap: ".65rem", padding: ".9rem", background: testResult.success ? "rgba(74,222,128,.06)" : "rgba(248,113,113,.06)", border: `1px solid ${testResult.success ? "rgba(74,222,128,.2)" : "rgba(248,113,113,.2)"}`, borderRadius: 2 }}>
                      {testResult.success ? <CheckCircle2 size={15} style={{ color: "#4ade80", flexShrink: 0 }} /> : <XCircle size={15} style={{ color: "#f87171", flexShrink: 0 }} />}
                      <div>
                        <div style={{ fontSize: ".78rem", fontWeight: 700, color: testResult.success ? "#4ade80" : "#f87171" }}>{testResult.success ? "Connection Verified!" : "Connection Failed"}</div>
                        <div style={{ fontSize: ".72rem", color: "rgba(245,240,232,.4)", marginTop: ".2rem" }}>{testResult.msg}</div>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* ══ TAB: LOGS ══════════════════════════════════════════════ */}
            {activeTab === "logs" && (
              <div className="sp-fade-in">
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "1.5rem" }}>
                  <div>
                    <h2 style={{ fontFamily: "'Clash Display', sans-serif", fontWeight: 500, fontSize: "1.5rem", letterSpacing: "-.02em", color: "var(--cream)" }}>Dispatch History</h2>
                    <p style={{ fontSize: ".78rem", color: "rgba(245,240,232,.3)", marginTop: ".3rem" }}>Full records of every campaign run, with per-recipient delivery status.</p>
                  </div>
                  {logs.length > 0 && (
                    <button className="sp-btn-dash sp-btn-ink" onClick={() => askConfirm("Clear all campaign history logs?", clearLogs)}>
                      <Trash2 size={13} /> Clear Logs
                    </button>
                  )}
                </div>

                {logs.length === 0 ? (
                  <div style={{ textAlign: "center", padding: "5rem 2rem", border: "2px dashed rgba(245,240,232,.06)", borderRadius: 4 }}>
                    <History size={40} style={{ color: "rgba(245,240,232,.08)", margin: "0 auto 1rem" }} />
                    <div style={{ fontFamily: "'Clash Display', sans-serif", fontSize: "1.1rem", color: "rgba(245,240,232,.25)" }}>No campaigns dispatched yet</div>
                    <div style={{ fontSize: ".78rem", color: "rgba(245,240,232,.15)", marginTop: ".4rem" }}>Campaign execution reports will appear here after you send.</div>
                  </div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: ".75rem" }} className="sp-stagger">
                    {logs.map(log => (
                      <div key={log._id} className="sp-card sp-lift" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "1rem" }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: ".6rem", marginBottom: ".4rem" }}>
                            <span style={{ fontSize: ".65rem", color: "rgba(245,240,232,.25)", letterSpacing: ".1em", textTransform: "uppercase" }}>{log.date}</span>
                            <span className={`sp-chip ${log.status === "Completed" ? "sp-chip-green" : "sp-chip-amber"}`}>{log.status}</span>
                          </div>
                          <div style={{ fontFamily: "'Clash Display', sans-serif", fontWeight: 500, fontSize: "1rem", letterSpacing: "-.01em", color: "var(--cream)", marginBottom: ".4rem", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 500 }}>{log.subject}</div>
                          <div style={{ display: "flex", gap: "1.25rem", fontSize: ".73rem" }}>
                            <span style={{ color: "rgba(245,240,232,.3)" }}>Total: <strong style={{ color: "rgba(245,240,232,.7)" }}>{log.total}</strong></span>
                            <span style={{ color: "rgba(245,240,232,.3)" }}>Delivered: <strong style={{ color: "#4ade80" }}>{log.success}</strong></span>
                            <span style={{ color: "rgba(245,240,232,.3)" }}>Failed: <strong style={{ color: "#f87171" }}>{log.failure}</strong></span>
                          </div>
                        </div>
                        <button className="sp-btn-dash sp-btn-ink" onClick={() => setSelectedLog(log)}>
                          <Eye size={13} style={{ color: "var(--rust)" }} /> View Report
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

          </div>
        </main>
      </div>

      {/* ── MODAL: SAVE TEMPLATE ── */}
      {showSaveTplModal && (
        <div className="sp-modal-bg" onClick={() => setShowSaveTplModal(false)}>
          <div className="sp-modal" onClick={e => e.stopPropagation()}>
            <div className="sp-modal-title">Save Template</div>
            <div className="sp-modal-sub">Name this draft to access it from the Template Library.</div>
            <div style={{ marginBottom: "1.25rem" }}>
              <label className="sp-label">Template Name</label>
              <input className="sp-input" value={templateName} onChange={e => setTemplateName(e.target.value)} placeholder="e.g. Summer Promo, Newsletter…" autoFocus />
            </div>
            <div style={{ display: "flex", gap: ".6rem", justifyContent: "flex-end" }}>
              <button className="sp-btn-dash sp-btn-ink" onClick={() => setShowSaveTplModal(false)}>Cancel</button>
              <button className="sp-btn-dash sp-btn-rust" onClick={saveTemplate}>Save Template</button>
            </div>
          </div>
        </div>
      )}

      {/* ── DRAWER: LOG DETAIL ── */}
      {selectedLog && (
        <>
          <div className="sp-drawer-bg" onClick={() => setSelectedLog(null)} />
          <div className="sp-drawer">
            <div style={{ padding: "1.75rem 1.5rem", borderBottom: "1px solid rgba(245,240,232,.06)", display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: "1rem" }}>
              <div>
                <div style={{ fontSize: ".65rem", fontWeight: 700, letterSpacing: ".15em", textTransform: "uppercase", color: "var(--rust)", marginBottom: ".4rem" }}>Campaign Report</div>
                <div style={{ fontFamily: "'Clash Display', sans-serif", fontWeight: 500, fontSize: "1.1rem", color: "var(--cream)", letterSpacing: "-.01em", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 340 }}>{selectedLog.subject}</div>
              </div>
              <button onClick={() => setSelectedLog(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "rgba(245,240,232,.3)", marginTop: 4 }}><X size={18} /></button>
            </div>

            <div style={{ padding: "1.25rem 1.5rem", borderBottom: "1px solid rgba(245,240,232,.06)", display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: ".75rem" }}>
              {[["Total",selectedLog.total,"rgba(245,240,232,.7)"],["Delivered",selectedLog.success,"#4ade80"],["Failed",selectedLog.failure,"#f87171"]].map(([l,v,c]) => (
                <div key={l} style={{ textAlign: "center", padding: ".85rem", background: "rgba(245,240,232,.03)", border: "1px solid rgba(245,240,232,.06)", borderRadius: 2 }}>
                  <div style={{ fontSize: ".6rem", fontWeight: 700, letterSpacing: ".15em", textTransform: "uppercase", color: "rgba(245,240,232,.25)", marginBottom: ".3rem" }}>{l}</div>
                  <div style={{ fontFamily: "'Clash Display', sans-serif", fontWeight: 600, fontSize: "1.8rem", color: c, letterSpacing: "-.04em" }}>{v}</div>
                </div>
              ))}
            </div>

            <div style={{ flex: 1, overflow: "auto", padding: "0 1.5rem 1.5rem" }}>
              <div style={{ paddingTop: "1rem", marginBottom: ".5rem" }}>
                <span style={{ fontSize: ".62rem", fontWeight: 700, letterSpacing: ".15em", textTransform: "uppercase", color: "rgba(245,240,232,.2)" }}>Delivery Details</span>
              </div>
              {selectedLog.details.map((d, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: ".7rem 0", borderBottom: "1px solid rgba(245,240,232,.04)" }}>
                  <span style={{ fontSize: ".8rem", color: "rgba(245,240,232,.6)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 280 }}>{d.email}</span>
                  <div style={{ display: "flex", alignItems: "center", gap: ".4rem", flexShrink: 0 }}>
                    {d.success
                      ? <><span style={{ fontSize: ".68rem", color: "#4ade80" }}>Delivered</span><CheckCircle2 size={14} style={{ color: "#4ade80" }} /></>
                      : <><span style={{ fontSize: ".68rem", color: "#f87171", maxWidth: 100, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{d.error}</span><XCircle size={14} style={{ color: "#f87171" }} /></>
                    }
                  </div>
                </div>
              ))}
            </div>

            <div style={{ padding: "1rem 1.5rem", borderTop: "1px solid rgba(245,240,232,.06)" }}>
              <button className="sp-btn-dash sp-btn-rust" style={{ width: "100%", justifyContent: "center" }} onClick={() => setSelectedLog(null)}>Close Report</button>
            </div>
          </div>
        </>
      )}

      {/* ── CAMPAIGN OVERLAY ── */}
      {sending.active && (
        <div className="sp-overlay">
          <div className="sp-overlay-card">
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
              <div>
                <div style={{ fontFamily: "'Clash Display', sans-serif", fontWeight: 500, fontSize: "1.25rem", letterSpacing: "-.02em", color: "var(--cream)", display: "flex", alignItems: "center", gap: ".6rem" }}>
                  <RefreshCw size={18} style={{ color: "var(--rust)", animation: "spSpin 1s linear infinite" }} />
                  {sending.paused ? "Campaign Paused" : "Dispatching…"}
                </div>
                <div style={{ fontSize: ".73rem", color: "rgba(245,240,232,.3)", marginTop: ".25rem", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 420 }}>Subject: {subject}</div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: ".5rem", padding: ".35rem .85rem", background: "rgba(245,240,232,.04)", border: "1px solid rgba(245,240,232,.06)", borderRadius: 20 }}>
                <span style={{ width: 8, height: 8, borderRadius: "50%", background: sending.paused ? "#fbbf24" : "#4ade80", boxShadow: `0 0 0 3px ${sending.paused ? "rgba(251,191,36,.2)" : "rgba(74,222,128,.2)"}`, animation: "spPulse 2s ease-in-out infinite" }} />
                <span style={{ fontSize: ".65rem", fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase", color: "rgba(245,240,232,.5)" }}>{sending.paused ? "Paused" : "Live"}</span>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: ".75rem", marginBottom: "1.25rem" }}>
              {[
                [`${sending.progress}/${sending.total}`, "Processed", "rgba(245,240,232,.7)"],
                [sending.success, "Delivered", "#4ade80"],
                [sending.fail, "Failed", "#f87171"],
              ].map(([v, l, c]) => (
                <div key={l} style={{ textAlign: "center", padding: "1rem", background: "rgba(245,240,232,.03)", border: "1px solid rgba(245,240,232,.06)", borderRadius: 2 }}>
                  <div style={{ fontSize: ".6rem", fontWeight: 700, letterSpacing: ".15em", textTransform: "uppercase", color: `${c}80`, marginBottom: ".3rem" }}>{l}</div>
                  <div style={{ fontFamily: "'Clash Display', sans-serif", fontWeight: 600, fontSize: "1.6rem", letterSpacing: "-.04em", color: c }}>{v}</div>
                </div>
              ))}
            </div>

            <div style={{ marginBottom: "1.25rem" }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: ".68rem", color: "rgba(245,240,232,.3)", marginBottom: ".4rem" }}>
                <span>Progress</span>
                <span style={{ color: "var(--rust)", fontWeight: 700 }}>{Math.round((sending.progress/sending.total)*100)}%</span>
              </div>
              <div className="sp-progress-bar">
                <div className="sp-progress-fill" style={{ width: `${(sending.progress/sending.total)*100}%` }} />
              </div>
            </div>

            <div style={{ marginBottom: "1.25rem" }}>
              <div style={{ fontSize: ".62rem", fontWeight: 700, letterSpacing: ".15em", textTransform: "uppercase", color: "rgba(245,240,232,.2)", marginBottom: ".5rem" }}>Console</div>
              <div className="sp-console">
                {consoleLogs.map((l, i) => (
                  <div key={i} style={{ color: l.includes("✓") ? "#4ade80" : l.includes("✗") ? "#f87171" : l.includes("⏸") || l.includes("▶") ? "#fbbf24" : "rgba(245,240,232,.4)" }}>{l}</div>
                ))}
                <div ref={consoleEnd} />
              </div>
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", borderTop: "1px solid rgba(245,240,232,.06)", paddingTop: "1rem" }}>
              <button className="sp-btn-dash sp-btn-danger" onClick={stopCampaign}><Square size={12} style={{ fill: "#f87171" }} /> Cancel</button>
              <button className={`sp-btn-dash ${sending.paused ? "sp-btn-rust" : ""}`} style={!sending.paused ? { background: "rgba(251,191,36,.1)", color: "#fbbf24", border: "1px solid rgba(251,191,36,.2)" } : {}} onClick={togglePause}>
                {sending.paused ? <><Play size={12} style={{ fill: "currentColor" }} /> Resume</> : <><Pause size={12} /> Pause</>}
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}