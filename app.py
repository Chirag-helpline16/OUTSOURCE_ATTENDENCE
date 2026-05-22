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
        :root {
            color-scheme: light;
            --att-primary: #2563eb;
            --att-primary-dark: #1d4ed8;
            --att-teal: #0f766e;
            --att-green: #16a34a;
            --att-amber: #d97706;
            --att-rose: #e11d48;
            --att-ink: #0f172a;
            --att-muted: #64748b;
            --att-line: #dbe3ef;
            --att-soft: #f6f8fb;
        }

        .stApp {
            background: #ffffff;
            color: var(--att-ink);
        }

        .block-container {
            padding-top: 1.6rem;
            padding-bottom: 2.5rem;
        }

        [data-testid="stSidebar"] {
            background:
                linear-gradient(180deg, #eff6ff 0%, #ffffff 34%, #f8fafc 100%);
            border-right: 1px solid var(--att-line);
            box-shadow: 12px 0 34px rgba(15, 23, 42, 0.06);
        }

        [data-testid="stSidebar"] * {
            color: var(--att-ink);
        }

        h1, h2, h3, h4, h5, h6,
        .stMarkdown, .stCaption, label, p {
            color: var(--att-ink);
        }

        h1 {
            padding-left: 0.85rem;
            border-left: 5px solid var(--att-primary);
            letter-spacing: 0;
        }

        h2, h3 {
            letter-spacing: 0;
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
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.06);
        }

        div[data-testid="stMetric"] {
            border-top: 4px solid var(--att-primary);
            padding: 0.85rem;
        }

        div[data-testid="stForm"] {
            padding: 1rem;
            border-top: 4px solid var(--att-teal);
        }

        .stDataFrame {
            border: 1px solid var(--att-line);
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.04);
        }

        [data-testid="stTabs"] button {
            color: var(--att-muted);
            font-weight: 700;
        }

        [data-testid="stTabs"] button[aria-selected="true"] {
            color: var(--att-primary);
            border-bottom-color: var(--att-primary);
        }

        div[role="radiogroup"] label {
            border: 1px solid var(--att-line);
            border-radius: 8px;
            background: #ffffff;
            padding: 0.35rem 0.55rem;
            margin-bottom: 0.35rem;
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
            border-color: var(--att-primary);
            box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.14);
        }

        .stButton > button,
        .stDownloadButton > button {
            border-radius: 8px;
            border: 1px solid var(--att-primary);
            background: var(--att-primary);
            color: #ffffff;
            font-weight: 600;
            box-shadow: 0 8px 18px rgba(37, 99, 235, 0.2);
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover {
            border-color: var(--att-primary-dark);
            background: var(--att-primary-dark);
            color: #ffffff;
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
            background: var(--att-rose) !important;
            border-color: var(--att-rose) !important;
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
        selected_page = st.radio(
            "Open page",
            options=list(PAGES.keys()),
            label_visibility="collapsed",
        )
        st.divider()
        st.caption("Outsource attendance system")

    PAGES[selected_page]()


if __name__ == "__main__":
    main()
