# flake.nix
{
  description = "A development environment for Office Add-ins";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";
    flake-utils.url = "github:numtide/flake-utils";
    rust-overlay = {
      url = "github:oxalica/rust-overlay";
      inputs = {
        nixpkgs.follows = "nixpkgs";
        flake-utils.follows = "flake-utils";
      };
    };
  };

  outputs = {
    self,
    nixpkgs,
    flake-utils,
    rust-overlay,
  }:
    flake-utils.lib.eachDefaultSystem (system: let
      # We'll use the unstable channel for newer nodejs versions
      system = "x86_64-linux";
      overlays = [(import rust-overlay)];
      pkgs = import nixpkgs {
        inherit system overlays;
      };
      rust =
        pkgs.rust-bin.stable.latest.default.override
        {
          extensions = ["rust-src"];
          targets = ["wasm32-unknown-unknown"];
        };
    in {
      # The 'nix develop' command will drop you into this shell
      devShells.default = with pkgs;
        mkShell {
          # These are the packages Nix will make available in our shell's PATH.
          buildInputs = [
            nodejs # A specific, recent version of Node.js
            yarn # Often useful for JS projects
            yo # The Yeoman scaffolding tool itself
            rust
          ];

          # This is a script that runs when you enter the shell.
          # We'll use it to configure npm to use a local directory
          # for "global" packages, instead of polluting the system.
          shellHook = ''
            echo "Welcome to the Office Add-in Dev Shell!"

            # Configure npm to install "global" packages into ./node_modules
            # instead of a system-wide directory.
            export NPM_CONFIG_PREFIX=$(pwd)/.npm-packages
            export PATH=$NPM_CONFIG_PREFIX/bin:$PATH

            echo "NPM global packages will be installed in $(pwd)/.npm-packages"
          '';
        };
    });
}
