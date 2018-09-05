import argparse
parser = argparse.ArgumentParser()
parser.add_argument('echo')
parser.add_argument('--verbosity', help='increase output verbosity')
args = parser.parse_args()
print(args.echo)
if args.verbosity:
    print('verbosity turned on')
    print(args.verbosity)
