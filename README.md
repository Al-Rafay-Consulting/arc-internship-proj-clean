# Gemini RAG Extraction Project

A Retrieval-Augmented Generation (RAG) based document extraction system that uses Google's Gemini models, Azure AI Search, and vector-based retrieval to extract structured information from documents and generate accurate, context-aware outputs.

## Overview

This project implements a RAG pipeline that combines document retrieval with Large Language Models (LLMs) to improve accuracy and reduce hallucination.

Instead of directly asking an LLM to answer from memory, the system:

1. Processes and indexes documents
2. Retrieves relevant information using semantic search
3. Provides retrieved context to Gemini
4. Generates structured responses based on the source documents

The project focuses on automated information extraction from multiple documents using customizable prompts.

---

## Features

- ✅ Retrieval-Augmented Generation (RAG) pipeline
- ✅ Integration with Google Gemini models
- ✅ Azure AI Search integration for document retrieval
- ✅ Semantic search-based context retrieval
- ✅ Automated extraction using predefined prompts
- ✅ Batch processing of multiple extraction tasks
- ✅ Structured output generation
- ✅ Environment-based API key management
- ✅ Exporting generated results into Word documents

---

## Architecture
