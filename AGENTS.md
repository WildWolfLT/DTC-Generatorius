# AGENTS: read this first

This repository is hosted only on DigitalOcean App Platform. Pushes to `main` redeploy https://dtc-generatorius-clom4.ondigitalocean.app. Do not start a local server.

## Required before you finish

Every update must be committed and pushed. No local change may remain uncommitted or unpushed.

1. `git add` the files you changed. Never commit secrets.
2. `git commit` with a message that says why.
3. `git push` to `origin` `main`.

Leave `git status` clean. Verify on the live URL after App Platform finishes deploying.

Serve CSS and JS from `/assets`. Requests to `/static` return HTTP 530 at the edge and never reach the app.

The same instructions are injected every session from `.cursor/rules/remote-hosting.mdc`.

## App spec

```yaml
name: dtc-generatorius
region: fra
features:
  - buildpack-stack=ubuntu-22
ingress:
  rules:
    - component:
        name: web
      match:
        path:
          prefix: /
services:
  - name: web
    environment_slug: python
    github:
      repo: WildWolfLT/DTC-Generatorius
      branch: main
      deploy_on_push: true
    http_port: 8080
    instance_count: 1
    instance_size_slug: apps-s-1vcpu-0.5gb
    run_command: gunicorn --worker-tmp-dir /dev/shm --bind 0.0.0.0:8080 --workers 2 --preload app:app
```
