# Plugins

> Extend Claude Code with custom commands, agents, hooks, Skills, and MCP servers through the plugin system.

Plugins let you extend Claude Code with custom functionality that can be shared across projects and teams. Install plugins from marketplaces to add pre-built commands, agents, hooks, Skills, and MCP servers, or create your own to automate your workflows.

## Quickstart

Let's create a simple greeting plugin to get you familiar with the plugin system. We'll build a working plugin that adds a custom command, test it locally, and understand the core concepts.

### Prerequisites

* Claude Code installed on your machine
* Basic familiarity with command-line tools

### Create your first plugin

1. **Create the marketplace structure**
   ```bash
   mkdir test-marketplace
   cd test-marketplace
   ```

2. **Create the plugin directory**
   ```bash
   mkdir my-first-plugin
   cd my-first-plugin
   ```

3. **Create the plugin manifest**
   ```bash
   # Create .claude-plugin/plugin.json
   mkdir .claude-plugin
   cat > .claude-plugin/plugin.json << 'EOF'
   {
     "name": "my-first-plugin",
     "description": "A simple greeting plugin to learn the basics",
     "version": "1.0.0",
     "author": {
       "name": "Your Name"
     }
   }
   EOF
   ```

4. **Add a custom command**
   ```bash
   # Create commands/hello.md
   mkdir commands
   cat > commands/hello.md << 'EOF'
   ---
   description: Greet the user with a personalized message
   ---

   # Hello Command

   Greet the user warmly and ask how you can help them today.
   EOF
   ```

5. **Create the marketplace manifest**
   ```bash
   # Create marketplace.json
   cd ..
   mkdir .claude-plugin
   cat > .claude-plugin/marketplace.json << 'EOF'
   {
     "name": "test-marketplace",
     "owner": {
       "name": "Test User"
     },
     "plugins": [
       {
         "name": "my-first-plugin",
         "source": "./my-first-plugin",
         "description": "My first test plugin"
       }
     ]
   }
   EOF
   ```

6. **Install and test your plugin**
   ```bash
   # Start Claude Code from parent directory
   cd ..
   claude

   # Add the test marketplace
   /plugin marketplace add ./test-marketplace

   # Install your plugin
   /plugin install my-first-plugin@test-marketplace

   # Try your new command
   /hello
   ```

### Plugin structure overview

Your plugin follows this basic structure:

```
my-first-plugin/
├── .claude-plugin/
│   └── plugin.json          # Plugin metadata
├── commands/                 # Custom slash commands (optional)
│   └── hello.md
├── agents/                   # Custom agents (optional)
│   └── helper.md
├── skills/                   # Agent Skills (optional)
│   └── my-skill/
│       └── SKILL.md
└── hooks/                    # Event handlers (optional)
    └── hooks.json
```

**Additional components you can add:**

* **Commands**: Create markdown files in `commands/` directory
* **Agents**: Create agent definitions in `agents/` directory
* **Skills**: Create `SKILL.md` files in `skills/` directory
* **Hooks**: Create `hooks/hooks.json` for event handling
* **MCP servers**: Create `.mcp.json` for external tool integration

---

## Install and manage plugins

### Add marketplaces

Marketplaces are catalogs of available plugins. Add them to discover and install plugins:

```shell
# Add a marketplace
/plugin marketplace add your-org/claude-plugins

# Browse available plugins
/plugin
```

### Install plugins

#### Via interactive menu (recommended for discovery)

```shell
# Open the plugin management interface
/plugin
```

Select "Browse Plugins" to see available options with descriptions, features, and installation options.

#### Via direct commands (for quick installation)

```shell
# Install a specific plugin
/plugin install formatter@your-org

# Enable a disabled plugin
/plugin enable plugin-name@marketplace-name

# Disable without uninstalling
/plugin disable plugin-name@marketplace-name

# Completely remove a plugin
/plugin uninstall plugin-name@marketplace-name
```

### Verify installation

After installing a plugin:

1. **Check available commands**: Run `/help` to see new commands
2. **Test plugin features**: Try the plugin's commands and features
3. **Review plugin details**: Use `/plugin` → "Manage Plugins" to see what the plugin provides

---

## Develop more complex plugins

### Add Skills to your plugin

Plugins can include Agent Skills to extend Claude's capabilities. Skills are model-invoked—Claude autonomously uses them based on the task context.

To add Skills to your plugin, create a `skills/` directory at your plugin root and add Skill folders with `SKILL.md` files. Plugin Skills are automatically available when the plugin is installed.

### Test your plugins locally

When developing plugins, use a local marketplace to test changes iteratively.

1. **Set up your development structure**
   ```bash
   mkdir dev-marketplace
   cd dev-marketplace
   mkdir my-plugin
   ```

   This creates:
   ```
   dev-marketplace/
   ├── .claude-plugin/marketplace.json  (you'll create this)
   └── my-plugin/                        (your plugin under development)
       ├── .claude-plugin/plugin.json
       ├── commands/
       ├── agents/
       └── hooks/
   ```

2. **Create the marketplace manifest**
   ```bash
   mkdir .claude-plugin
   cat > .claude-plugin/marketplace.json << 'EOF'
   {
     "name": "dev-marketplace",
     "owner": {
       "name": "Developer"
     },
     "plugins": [
       {
         "name": "my-plugin",
         "source": "./my-plugin",
         "description": "Plugin under development"
       }
     ]
   }
   EOF
   ```

3. **Install and test**
   ```bash
   cd ..
   claude

   /plugin marketplace add ./dev-marketplace
   /plugin install my-plugin@dev-marketplace
   ```

4. **Iterate on your plugin**
   ```shell
   # Uninstall the current version
   /plugin uninstall my-plugin@dev-marketplace

   # Reinstall to test changes
   /plugin install my-plugin@dev-marketplace
   ```

### Debug plugin issues

If your plugin isn't working as expected:

1. **Check the structure**: Ensure your directories are at the plugin root, not inside `.claude-plugin/`
2. **Test components individually**: Check each command, agent, and hook separately
3. **Use validation and debugging tools**: See Plugins reference for CLI commands

### Share your plugins

When your plugin is ready to share:

1. **Add documentation**: Include a README.md with installation and usage instructions
2. **Version your plugin**: Use semantic versioning in your `plugin.json`
3. **Create or use a marketplace**: Distribute through plugin marketplaces for installation
4. **Test with others**: Have team members test the plugin before wider distribution

---

## See also

* Plugin marketplaces - Creating and managing plugin catalogs
* Slash commands - Understanding custom commands
* Subagents - Creating and using specialized agents
* Agent Skills - Extend Claude's capabilities
* Hooks - Automating workflows with event handlers
* MCP - Connecting to external tools and services
