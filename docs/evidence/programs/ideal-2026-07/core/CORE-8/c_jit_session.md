# Bounded embedded JIT sessions

2026-10-31; `bd-59co.2.10.2` remains open. The facade retains compiled code and per-session ExecState, globals, Static and Err; budgets restart per invocation, separate sessions isolate state and reset retains code. Range callbacks use the synchronous borrowed host context. [Run 3 evidence](udf-session-embedding.md) records the tested delivery and limitations.

This does not establish general project-class disposal, descriptor generations, recursion/reentry, thread/apartment transit or Windows certification. Shared CORE-5 ownership and broader CORE-8 work remain required before whole-row completion.
