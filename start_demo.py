#!/usr/bin/env python3
# Copyright 2024 PPT Generate Project
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""Quick start script for the PPT Generate MCP Server."""

import os
import sys
from pathlib import Path

def main():
    """Start the MCP server."""
    print("🚀 Starting PPT Generate MCP Server...")

    # Set PYTHONPATH to include the project root
    project_root = str(Path(__file__).parent.absolute())
    if "PYTHONPATH" in os.environ:
        os.environ["PYTHONPATH"] = project_root + os.pathsep + os.environ["PYTHONPATH"]
    else:
        os.environ["PYTHONPATH"] = project_root

    # Import and run the MCP server
    try:
        from src.mcp.server import main as mcp_main
        mcp_main()
    except ImportError as e:
        print(f"❌ Error: Could not import MCP server. Make sure dependencies are installed. {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n👋 MCP Server stopped by user")
    except Exception as e:
        print(f"❌ Failed to start MCP Server: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
