# ADR-0005: PathPolicy の安全境界

## 状態

`accepted`

## 背景

MCP tool は外部クライアントから渡された file path を操作する。どの path を許可するか、出力 path をどう解決するか、path validation をどこに置くかについて、明確な安全境界が必要である。

## 決定

- `PathPolicy` を MCP の file access と output location rule を検証する中心境界とする。
- caller は tool call site ごとに path check を複製せず、`PathPolicy` を意識した helper を通じて path を正規化・検証する。
- safety check は任意の hygiene ではなく、tool execution contract の一部として扱う。

## 影響

- 新しい MCP file-manipulating tool は `PathPolicy` と統合しなければならない。
- MCP handler での直接 filesystem access は、テスト済み boundary helper で正当化されない限り疑うべきである。
- path behavior を変える場合は、allow、deny、normalization の test が必要になる。

## 根拠

- Tests: `tests/mcp/test_path_policy.py`, `tests/mcp/test_validate_input.py`, `tests/mcp/shared/test_output_path.py`
- Code: `src/exstruct/mcp/io.py`, `src/exstruct/mcp/server.py`, `src/exstruct/mcp/render_runner.py`
- Related specs: `docs/mcp.md`

## Supersedes

- None

## Superseded by

- None
