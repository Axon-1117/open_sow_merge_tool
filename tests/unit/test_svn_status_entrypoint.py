from __future__ import annotations

import json

from sow_merge_tool import svn_status_provider as provider


def test_internal_status_entrypoint_accepts_remote_flag(tmp_path, monkeypatch):
    output = tmp_path / "status.json"
    seen = []

    def fake_query(path, *, remote=False):
        seen.append((path, remote))
        return [provider.SvnStatusRecord(path=path, node_kind="dir", node_status="normal")]

    monkeypatch.setattr(provider, "query_tortoise_status_in_child", fake_query)

    result = provider.internal_status_entrypoint(
        [str(tmp_path), str(output), "--remote"]
    )

    assert result == 0
    assert seen == [(str(tmp_path), True)]
    payload = json.loads(output.read_text(encoding="utf-8"))
    assert payload["ok"] is True
    assert payload["items"][0]["node_status"] == "normal"


def test_internal_status_entrypoint_accepts_local_mode(tmp_path, monkeypatch):
    output = tmp_path / "status.json"
    seen = []

    def fake_query(path, *, remote=False):
        seen.append(remote)
        return []

    monkeypatch.setattr(provider, "query_tortoise_status_in_child", fake_query)

    result = provider.internal_status_entrypoint([str(tmp_path), str(output)])

    assert result == 0
    assert seen == [False]
    assert json.loads(output.read_text(encoding="utf-8"))["items"] == []


def test_cli_status_uses_against_revision_for_remote_freshness():
    xml = """
    <status>
      <target path="C:/wc">
        <entry path="Quest.xlsx">
          <wc-status item="normal" props="normal" revision="41226">
            <commit revision="41220" />
          </wc-status>
        </entry>
        <against revision="41226" />
      </target>
    </status>
    """

    records = provider._parse_cli_status(xml, r"C:\wc")

    assert len(records) == 1
    assert records[0].revision == 41226
    assert records[0].repository_revision == 41226
    assert records[0].remote_revision == 41226
