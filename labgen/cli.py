"""Командный интерфейс labgen."""

from __future__ import annotations

import shutil
from pathlib import Path

import click

from . import __version__
from .config import Config, find_config
from .render import render_markdown

ASSETS = Path(__file__).parent / "assets"


def _echo_warnings(warnings: list[str]) -> None:
    for warning in warnings:
        click.secho(f"  ! {warning}", fg="yellow")


@click.group(context_settings={"help_option_names": ["-h", "--help"]})
@click.version_option(__version__, "-V", "--version", message="labgen %(version)s")
def main() -> None:
    """Отчёты по лабораторным: Markdown на входе, DOCX по ГОСТ 7.32-2017 на выходе."""


@main.command()
@click.argument("directory", type=click.Path(file_okay=False, path_type=Path), default=".")
@click.option("--force", is_flag=True, help="Перезаписать существующие файлы.")
def init(directory: Path, force: bool) -> None:
    """Создать labgen.yaml и заготовку отчёта в DIRECTORY."""
    directory.mkdir(parents=True, exist_ok=True)
    created: list[Path] = []
    for name in ("labgen.yaml", "report.md"):
        target = directory / name
        if target.exists() and not force:
            click.secho(f"  = {name} уже существует, пропускаю (--force перезапишет)", fg="yellow")
            continue
        shutil.copyfile(ASSETS / name, target)
        created.append(target)

    images = directory / "images"
    images.mkdir(exist_ok=True)

    for path in created:
        click.secho(f"  + {path}", fg="green")
    click.echo("\nЗаполните labgen.yaml и запустите: labgen build report.md")


@main.command()
@click.argument("source", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("-o", "--output", type=click.Path(dir_okay=False, path_type=Path),
              help="Куда сохранить .docx (по умолчанию — рядом с исходником).")
@click.option("-c", "--config", "config_path", help="Путь к labgen.yaml.",
              type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--title/--no-title", "title_page", default=None,
              help="Печатать титульный лист (по умолчанию — из конфигурации).")
@click.option("--toc/--no-toc", default=None,
              help="Вставлять содержание (по умолчанию — из конфигурации).")
@click.option("--strict", is_flag=True, help="Считать предупреждения ошибками.")
def build(source: Path, output: Path | None, config_path: Path | None,
          title_page: bool | None, toc: bool | None, strict: bool) -> None:
    """Собрать отчёт из markdown-файла SOURCE."""
    resolved_config = find_config(config_path, source)
    config = Config.load(resolved_config)
    if resolved_config:
        click.echo(f"  конфигурация: {resolved_config}")
    else:
        click.secho("  конфигурация не найдена, использую значения по умолчанию", fg="yellow")

    if title_page is None:
        title_page = bool(config.title_page.get("enabled", True))

    markdown_text = source.read_text(encoding="utf-8")
    document, result = render_markdown(
        markdown_text, config, source.parent, title_page=title_page, toc=toc
    )

    destination = output or source.with_suffix(".docx")
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        document.save(destination)
    except PermissionError as error:
        raise click.ClickException(
            f"{destination} занят другой программой — закройте файл в Word и повторите."
        ) from error

    click.secho(f"  + {destination}", fg="green")
    click.echo(
        f"  рисунков: {result.figures}, таблиц: {result.tables}, листингов: {result.listings}"
    )
    _echo_warnings(result.warnings)
    if strict and result.warnings:
        raise click.ClickException("сборка остановлена: есть предупреждения (--strict)")


if __name__ == "__main__":
    main()
