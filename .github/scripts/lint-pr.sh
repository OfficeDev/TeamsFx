#!/bin/bash
VAR=$(git diff --diff-filter=MARC $1...HEAD --name-only --relative -- .| grep -E '.js$|.ts$|.jsx$|.tsx$' | xargs)
echo $VAR
if [ ! -z "$VAR" ]
then 
    pnpm dlx eslint --quiet --fix $VAR
fi