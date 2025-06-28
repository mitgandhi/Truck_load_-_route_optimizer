import argparse
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description="Truck load optimizer")
    parser.add_argument(
        "--algorithm",
        choices=["ffd", "bestfit"],
        default="ffd",
        help="Optimization algorithm to use",
    )
    parser.add_argument(
        "--file",
        default="PRE_ORDER.xlsx",
        help="Path to input Excel file",
    )
    args = parser.parse_args()

    if args.algorithm == "ffd":
        subprocess.run([sys.executable, "FFD.py", args.file])
    else:
        from best_fit import main as best_main

        best_main(args.file)


if __name__ == "__main__":
    main()
