from __future__ import annotations

from pathlib import Path

import docx
from click.testing import CliRunner

from labgen.cli import main


def test_init_creates_scaffolding(tmp_path: Path):
    result = CliRunner().invoke(main, ["init", str(tmp_path)])
    assert result.exit_code == 0, result.output
    assert (tmp_path / "labgen.yaml").exists()
    assert (tmp_path / "report.md").exists()
    assert (tmp_path / "images").is_dir()


def test_init_does_not_overwrite_without_force(tmp_path: Path):
    runner = CliRunner()
    runner.invoke(main, ["init", str(tmp_path)])
    (tmp_path / "report.md").write_text("мой текст", encoding="utf-8")

    runner.invoke(main, ["init", str(tmp_path)])
    assert (tmp_path / "report.md").read_text(encoding="utf-8") == "мой текст"

    runner.invoke(main, ["init", str(tmp_path), "--force"])
    assert (tmp_path / "report.md").read_text(encoding="utf-8") != "мой текст"


def test_build_produces_document_next_to_source(tmp_path: Path):
    runner = CliRunner()
    runner.invoke(main, ["init", str(tmp_path)])
    result = runner.invoke(main, ["build", str(tmp_path / "report.md")])

    assert result.exit_code == 0, result.output
    output = tmp_path / "report.docx"
    assert output.exists()
    texts = [p.text for p in docx.Document(output).paragraphs if p.text.strip()]
    assert "Введение" in texts
    assert any(text.startswith("Таблица 1") for text in texts)


def test_build_honours_output_option(tmp_path: Path):
    runner = CliRunner()
    runner.invoke(main, ["init", str(tmp_path)])
    target = tmp_path / "out" / "лаба3.docx"
    result = runner.invoke(
        main, ["build", str(tmp_path / "report.md"), "-o", str(target), "--no-toc"]
    )
    assert result.exit_code == 0, result.output
    assert target.exists()


def test_strict_fails_on_warnings(tmp_path: Path):
    source = tmp_path / "report.md"
    source.write_text("![Нет файла](images/absent.png)\n", encoding="utf-8")
    result = CliRunner().invoke(main, ["build", str(source), "--strict", "--no-title"])
    assert result.exit_code != 0
    assert "предупреждения" in result.output
