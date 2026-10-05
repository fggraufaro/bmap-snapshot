"""Engagement Builder: guided intake -> reviewed SOW draft (DOCX) -> sign-off log.

Clause text, module scope text and prices live in private Supabase tables
(public.engagement_*), never in this repository, which is public. The code
here is the generic engine: field validation, calculated dates and fees,
assembly, assertions, DOCX output and the Claude-backed chat and narrative
drafting with guardrails.
"""
