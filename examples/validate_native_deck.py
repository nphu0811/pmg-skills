"""Small real API example; engineering evidence stays internal."""
import argparse
from pmg import PMGOrchestrator, ValidationStatus, build_environmental_protection_deck


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('-o', '--output', required=True)
    parser.add_argument('--worker', default='auto', choices=['auto', 'local', 'remote', 'static'])
    args = parser.parse_args()
    manifest = PMGOrchestrator().run_pipeline(
        output_pptx_path=args.output,
        builder_fn=lambda path: build_environmental_protection_deck(path, num_slides=4),
        worker_mode=args.worker)
    print(f'{manifest.status.value}: {manifest.delivered_pptx or manifest.qa_dir}')
    return 0 if manifest.status == ValidationStatus.PASS else 1


if __name__ == '__main__':
    raise SystemExit(main())
