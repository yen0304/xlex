---
name: /commit
id: git-commit
category: Git
description: Create a conventional commit with proper formatting following industry standards.
---

**Conventional Commits Format**

Follow the [Conventional Commits](https://www.conventionalcommits.org/) specification:

```
<type>[optional scope]: <description>

[optional body]

[optional footer(s)]
```

**Allowed Types**
- `feat`: A new feature
- `fix`: A bug fix
- `docs`: Documentation only changes
- `style`: Changes that do not affect the meaning of the code (white-space, formatting, etc.)
- `refactor`: A code change that neither fixes a bug nor adds a feature
- `perf`: A code change that improves performance
- `test`: Adding missing tests or correcting existing tests
- `build`: Changes that affect the build system or external dependencies
- `ci`: Changes to CI configuration files and scripts
- `chore`: Other changes that don't modify src or test files
- `revert`: Reverts a previous commit

**Scope (optional)**
- A noun describing the section of the codebase (e.g., `api`, `frontend`, `auth`, `db`)

**Breaking Changes**
- Add `!` after type/scope for breaking changes: `feat!:` or `feat(api)!:`
- Or add `BREAKING CHANGE:` in the footer

**Steps**
1. Run `git status` to check current changes
2. Run `git diff --cached --stat` to see staged changes (if any)
3. Run `git diff --stat` to see unstaged changes
4. Analyze the changes and determine the appropriate type and scope
5. Stage relevant files with `git add` (group related changes if needed)
6. Write a clear, concise commit message following the format above
7. Commit using HEREDOC format for proper formatting:

```bash
git commit -m "$(cat <<'EOF'
<type>(<scope>): <short description>

<body - explain what and why, not how>

EOF
)"
```

**Staged Commits (分段 Commit)**

When there are multiple unrelated changes, split them into separate commits by feature/purpose:

1. **Analyze all changes first**
   - Run `git diff` to review all unstaged changes
   - Run `git diff --name-only` to list all modified files
   - Group files by feature, bug fix, or logical unit

2. **Identify logical groups**
   - Group by feature: files that implement the same feature together
   - Group by type: separate `feat`, `fix`, `refactor`, `docs`, `test` changes
   - Group by scope: changes to `api`, `frontend`, `backend`, etc.

3. **Stage and commit each group separately**
   ```bash
   # Group 1: Feature A
   git add src/feature-a/*.ts tests/feature-a/*.test.ts
   git commit -m "feat(feature-a): add new feature A"

   # Group 2: Bug fix
   git add src/utils/helper.ts
   git commit -m "fix(utils): resolve null pointer in helper"

   # Group 3: Documentation
   git add docs/*.md README.md
   git commit -m "docs: update API documentation"
   ```

4. **Interactive staging for partial file changes**
   - Use `git add -p <file>` to stage specific hunks within a file
   - Review each hunk and choose: `y` (stage), `n` (skip), `s` (split smaller)

5. **Verify before each commit**
   - Run `git diff --cached` to review what will be committed
   - Ensure the staged changes are cohesive and belong together

**Staged Commits Example Workflow**

Suppose you have changes across multiple files:
```
modified:   src/api/auth.ts      (new login feature)
modified:   src/api/user.ts      (new login feature)
modified:   src/utils/format.ts  (bug fix)
modified:   tests/auth.test.ts   (new login feature)
modified:   docs/api.md          (documentation)
```

Split into 3 commits:
```bash
# Commit 1: New feature
git add src/api/auth.ts src/api/user.ts tests/auth.test.ts
git commit -m "feat(auth): implement OAuth2 login flow"

# Commit 2: Bug fix
git add src/utils/format.ts
git commit -m "fix(utils): handle empty string in format function"

# Commit 3: Documentation
git add docs/api.md
git commit -m "docs(api): add authentication endpoint examples"
```

**Examples**

Simple commit:
```
feat(auth): add OAuth2 login support
```

With body:
```
fix(api): resolve race condition in user creation

The previous implementation could create duplicate users when
concurrent requests arrived. Added database-level unique constraint
and proper error handling.
```

Breaking change:
```
feat(api)!: change authentication header format

BREAKING CHANGE: Authorization header now requires 'Bearer ' prefix.
Clients must update their authentication logic accordingly.
```

**Best Practices**
- Use imperative mood: "add" not "added" or "adds"
- Don't capitalize the first letter of description
- No period at the end of the subject line
- Keep subject line under 50 characters
- Wrap body at 72 characters
- Use the body to explain *what* and *why*, not *how*
- **NO emojis** - Do not use any emoji or emoticon symbols in commit messages