#!/bin/bash
if [ $1 == 'templates' ]; then
    if [[ $SkipSyncup == *"template"* ]]; then
        echo "skip sync up templates version with sdk version"
    elif [[ -z "$(git diff -- ../../templates)" ]]; then
        echo "need bump up templates version since templates don not bump up by self"
        node ../../.github/scripts/sdk-sync-up-version.js sdk yes;
    else 
        echo "no need to bump up templates version"
        node ../../.github/scripts/sdk-sync-up-version.js sdk
    fi
    git add ../../templates
elif [ $1 == 'fx-core' ]; then
    if [[ -z "$(git diff -- ../fx-core)" ]]; then
        echo "need bump up fx-core version since fx-core don not bump up by self"
        node ../../.github/scripts/sync-up-dotnet-ver.js yes;
    else 
        echo "no need to bump up templates version"
        node ../../.github/scripts/sync-up-dotnet-ver.js
    fi
    git add ../fx-core
elif [ $1 == 'function-extension' ]; then   
    if [[ -z "$(git diff -- ../../templates)" ]]; then
        echo "need to bump up templates version since templates do not bump up by themselves"
        node ../../.github/scripts/sync-up-dotnet-ver.js yes
    else 
        echo "no need to bump up templates version"
        node ../../.github/scripts/sync-up-dotnet-ver.js
    fi
    git add ../../templates
elif [ $1 == 'core-template' ]; then
    echo "need to bump up templates' fallback version in fx-core"
    node ../.github/scripts/fxcore-sync-up-version.js
    git add ../packages/fx-core
    echo "sync up all sub templates deps"
    node ../.github/scripts/sync-templates.js
    git add ../templates
elif [ $1 == 'template-adaptive-card' ]; then
    if [[ $SkipSyncup == *"template"* ]]; then
        echo "skip sync up templates version with adaptive-card version"
    elif [[ -z "$(git diff -- ../../templates)" ]]; then
        echo "need bump up templates version since templates don not bump up by self"
        node ../../.github/scripts/sdk-sync-up-version.js adaptivecards-tools-sdk yes;
    else 
        echo "no need to bump up templates version"
        node ../../.github/scripts/sdk-sync-up-version.js adaptivecards-tools-sdk
    fi
    git add ../../templates
elif [ $1 == 'template-sync' ]; then
    subTempUpdate=$(node ../.github/scripts/sync-templates.js)
    git add .
fi