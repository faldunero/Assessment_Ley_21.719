#!/bin/bash
# Publica frontend/index.html en la rama gh-pages (GitHub Pages)
set -e
REPO="/Users/felipealdunate/Desktop/Desarrollo/ley21719-assessment"
cd "$REPO"

if [ -n "$(git status --porcelain)" ]; then
  echo "Hay cambios sin confirmar. Haz commit primero:"
  git status --short
  exit 1
fi

git fetch origin gh-pages
TMP=$(mktemp -d)
git worktree add "$TMP/pages" origin/gh-pages
cp frontend/index.html "$TMP/pages/index.html"
cd "$TMP/pages"
if git diff --quiet; then
  echo "La página ya está al día."
else
  git add index.html
  git commit -m "Publicar front ($(date +%Y-%m-%d))"
  git push origin HEAD:gh-pages
  echo "Publicado: https://faldunero.github.io/Assessment_Ley_21.719/"
fi
cd "$REPO"
git worktree remove --force "$TMP/pages"
rm -rf "$TMP"
