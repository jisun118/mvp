#!/bin/bash

# Streamlit 앱을 특정 포트에서 실행
python -m streamlit run EmailAnalyzer.py server.port=8081 --server.address=0.0.0.0 --server.enableCORS=false --server.enableXsrfProtection=false
