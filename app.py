"""Standalone Streamlit app for outsource attendance."""

from __future__ import annotations

import streamlit as st

from src.outsource_attendance import (
    render_attendance_admin_page,
    render_attendance_observer_page,
    render_outsource_login_page,
)


PAGES = {
    "Outsource Login": render_outsource_login_page,
    "Observer Desk": render_attendance_observer_page,
    "Admin Panel": render_attendance_admin_page,
}


def apply_professional_theme() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Michroma&family=Space+Mono:wght@400;700&display=swap');
        @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@20..48,400,0,0&display=block');

        :root {
            color-scheme: light;
            --att-red: #F04348;
            --att-amber: #FF9F0A;
            --att-blue: #2F80FF;
            --att-green: #21C16B;
            --att-ink: #090C12;
            --att-muted: #617086;
            --att-faint: #91A0B8;
            --att-line: #D9E1EC;
            --att-line-strong: #B9C6D8;
            --att-soft: #F6F8FB;
            --att-panel: #FFFFFF;
            --att-font: 'Space Mono', ui-monospace, SFMono-Regular, Consolas, monospace;
            --att-display: 'Michroma', 'Space Mono', ui-monospace, monospace;
        }

        .stApp {
            background: #ffffff;
            color: var(--att-ink);
        }

        html, body, .stApp, [class*="css"] {
            font-family: var(--att-font) !important;
        }

        body * {
            font-family: var(--att-font) !important;
            letter-spacing: 0 !important;
        }

        .material-icons,
        .material-icons-outlined,
        .material-symbols-rounded,
        .material-symbols-outlined,
        [data-testid="stIconMaterial"],
        span[class*="material-icons"],
        span[class*="material-symbols"],
        button[data-testid="stBaseButton-headerNoPadding"] span {
            font-family: 'Material Symbols Rounded', 'Material Icons', 'Material Icons Outlined' !important;
            font-weight: normal !important;
            font-style: normal !important;
            font-size: 20px !important;
            line-height: 1 !important;
            letter-spacing: normal !important;
            text-transform: none !important;
            white-space: nowrap !important;
            word-wrap: normal !important;
            direction: ltr !important;
            -webkit-font-feature-settings: 'liga' !important;
            -webkit-font-smoothing: antialiased !important;
            font-feature-settings: 'liga' !important;
        }

        button[data-testid="stBaseButton-headerNoPadding"] {
            width: 32px !important;
            height: 32px !important;
            min-height: 32px !important;
            padding: 4px !important;
            overflow: hidden !important;
            border-radius: 8px !important;
            color: var(--att-ink) !important;
        }

        .block-container {
            padding-top: 1.6rem;
            padding-bottom: 2.5rem;
        }

        [data-testid="stSidebar"] {
            background:
                linear-gradient(180deg, #ffffff 0%, #F8FAFC 42%, #ffffff 100%);
            border-right: 1px solid var(--att-line);
            box-shadow: 12px 0 34px rgba(9, 12, 18, 0.06);
        }

        [data-testid="stSidebar"] * {
            color: var(--att-ink);
        }

        h1, h2, h3, h4, h5, h6,
        .stMarkdown, .stCaption, label, p {
            color: var(--att-ink);
        }

        h1 {
            padding-left: 0;
            border-left: 0;
            font-family: var(--att-display) !important;
            font-weight: 400;
            letter-spacing: 0;
            position: relative;
        }

        h1::after {
            content: "";
            display: block;
            width: min(320px, 100%);
            height: 4px;
            margin-top: 0.65rem;
            border-radius: 999px;
            background: linear-gradient(90deg, var(--att-red), var(--att-amber), var(--att-blue), var(--att-green));
        }

        [data-testid="stSidebar"] h1::after {
            display: none;
        }

        h2, h3 {
            letter-spacing: 0;
            font-weight: 700;
        }

        .attendance-accent-line {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 4px;
            margin: 0.2rem 0 1rem;
        }

        .attendance-accent-line span {
            height: 4px;
            border-radius: 999px;
            display: block;
        }

        .attendance-accent-line span:nth-child(1) {
            background: var(--att-red);
        }

        .attendance-accent-line span:nth-child(2) {
            background: var(--att-amber);
        }

        .attendance-accent-line span:nth-child(3) {
            background: var(--att-blue);
        }

        .attendance-accent-line span:nth-child(4) {
            background: var(--att-green);
        }

        div[data-testid="stCaptionContainer"] p,
        .stCaption p {
            color: var(--att-muted);
        }

        div[data-testid="stMetric"],
        div[data-testid="stForm"],
        div[data-testid="stExpander"] {
            background: #ffffff;
            border: 1px solid var(--att-line);
            border-radius: 8px;
            box-shadow: 0 10px 24px rgba(9, 12, 18, 0.06);
        }

        div[data-testid="stMetric"] {
            border-top: 4px solid var(--att-red);
            padding: 0.85rem;
        }

        div[data-testid="stMetric"]:nth-of-type(2n) {
            border-top-color: var(--att-amber);
        }

        div[data-testid="stMetric"]:nth-of-type(3n) {
            border-top-color: var(--att-blue);
        }

        div[data-testid="stMetric"]:nth-of-type(4n) {
            border-top-color: var(--att-green);
        }

        div[data-testid="stForm"] {
            padding: 1rem;
            border-top: 4px solid var(--att-blue);
        }

        .stDataFrame {
            border: 1px solid var(--att-line);
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 8px 22px rgba(9, 12, 18, 0.04);
        }

        [data-testid="stTabs"] button {
            color: var(--att-muted);
            font-weight: 700;
        }

        [data-testid="stTabs"] button[aria-selected="true"] {
            color: var(--att-red);
            border-bottom-color: var(--att-red);
        }

        div[role="radiogroup"] label {
            border: 1px solid var(--att-line);
            border-radius: 8px;
            background: #ffffff;
            padding: 0.35rem 0.55rem;
            margin-bottom: 0.35rem;
        }

        [data-testid="stSidebar"] .st-key-main_navigation div[role="radiogroup"] {
            display: grid;
            gap: 0.65rem;
            width: 100%;
        }

        [data-testid="stSidebar"] .st-key-main_navigation div[role="radiogroup"] label {
            width: 100%;
            min-height: 48px;
            box-sizing: border-box;
            display: flex;
            align-items: center;
            gap: 0.55rem;
            padding: 0.75rem 0.9rem;
            margin: 0;
            border: 1px solid var(--att-line-strong);
            border-left: 5px solid var(--att-red);
            border-radius: 8px;
            background: #ffffff;
            box-shadow: 0 6px 14px rgba(9, 12, 18, 0.05);
            transition:
                transform 160ms ease,
                box-shadow 160ms ease,
                border-color 160ms ease,
                background 160ms ease;
        }

        [data-testid="stSidebar"] .st-key-main_navigation div[role="radiogroup"] label:nth-of-type(2) {
            border-left-color: var(--att-amber);
        }

        [data-testid="stSidebar"] .st-key-main_navigation div[role="radiogroup"] label:nth-of-type(3) {
            border-left-color: var(--att-blue);
        }

        [data-testid="stSidebar"] .st-key-main_navigation div[role="radiogroup"] label:nth-of-type(4) {
            border-left-color: var(--att-green);
        }

        [data-testid="stSidebar"] .st-key-main_navigation div[role="radiogroup"] label:hover {
            transform: translateX(4px);
            border-color: var(--att-blue);
            background: linear-gradient(90deg, rgba(47, 128, 255, 0.10), #ffffff 72%);
            box-shadow: 0 10px 20px rgba(47, 128, 255, 0.12);
        }

        [data-testid="stSidebar"] .st-key-main_navigation div[role="radiogroup"] label:has(input:checked) {
            border-color: var(--att-red);
            background: linear-gradient(90deg, rgba(240, 67, 72, 0.12), #ffffff 72%);
            box-shadow: 0 12px 24px rgba(240, 67, 72, 0.16);
            transform: translateX(2px);
        }

        [data-testid="stSidebar"] .st-key-main_navigation div[role="radiogroup"] label p {
            font-weight: 700;
            color: var(--att-ink);
            white-space: nowrap;
        }

        div[data-baseweb="input"] > div,
        div[data-baseweb="select"] > div,
        textarea {
            background: #ffffff;
            border-color: #cbd5e1;
            color: var(--att-ink);
        }

        div[data-baseweb="input"] > div:focus-within,
        div[data-baseweb="select"] > div:focus-within,
        textarea:focus {
            border-color: var(--att-blue);
            box-shadow: 0 0 0 3px rgba(47, 128, 255, 0.14);
        }

        .stButton > button,
        .stDownloadButton > button {
            border-radius: 8px;
            border: 1px solid var(--att-red);
            background: var(--att-red);
            color: var(--att-ink);
            font-weight: 600;
            box-shadow: 0 8px 18px rgba(240, 67, 72, 0.18);
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover {
            border-color: var(--att-amber);
            background: var(--att-amber);
            color: var(--att-ink);
        }

        .st-key-admin_decision_accept button,
        .st-key-observer_decision_accept button {
            background: var(--att-green) !important;
            border-color: var(--att-green) !important;
            color: #ffffff !important;
            box-shadow: 0 8px 18px rgba(22, 163, 74, 0.22) !important;
        }

        .st-key-admin_decision_accept button:hover,
        .st-key-observer_decision_accept button:hover {
            background: #15803d !important;
            border-color: #15803d !important;
            color: #ffffff !important;
        }

        .st-key-admin_decision_reject button,
        .st-key-observer_decision_reject button {
            background: var(--att-red) !important;
            border-color: var(--att-red) !important;
            color: #ffffff !important;
            box-shadow: 0 8px 18px rgba(225, 29, 72, 0.22) !important;
        }

        .st-key-admin_decision_reject button:hover,
        .st-key-observer_decision_reject button:hover {
            background: #be123c !important;
            border-color: #be123c !important;
            color: #ffffff !important;
        }

        .st-key-admin_logout button,
        .st-key-observer_logout button,
        .st-key-outsource_logout button {
            min-height: 2.5rem;
            padding: 0.55rem 0.8rem;
            background: #ffffff !important;
            border-color: #cbd5e1 !important;
            color: var(--att-ink) !important;
            box-shadow: none !important;
        }

        .st-key-admin_logout button:hover,
        .st-key-observer_logout button:hover,
        .st-key-outsource_logout button:hover {
            background: #fee2e2 !important;
            border-color: #fecaca !important;
            color: #991b1b !important;
        }

        .stAlert {
            border-radius: 8px;
        }

        hr {
            border-color: var(--att-line);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="Outsource Attendance",
        page_icon="ATT",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_professional_theme()

    with st.sidebar:
        st.title("Attendance")
        st.markdown(
            '<div class="attendance-accent-line"><span></span><span></span><span></span><span></span></div>',
            unsafe_allow_html=True,
        )
        selected_page = st.radio(
            "Open page",
            options=list(PAGES.keys()),
            label_visibility="collapsed",
            key="main_navigation",
        )
        st.divider()
        st.caption("Outsource attendance system")

    PAGES[selected_page]()


if __name__ == "__main__":
    main()
