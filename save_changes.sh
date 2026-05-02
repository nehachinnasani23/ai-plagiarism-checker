#!/bin/zsh

if [ -z "$1" ]; then
  echo 'Usage: ./save_changes.sh "your commit message"'
  exit 1
fi

git status --short
git add .
git commit -m "$1"
