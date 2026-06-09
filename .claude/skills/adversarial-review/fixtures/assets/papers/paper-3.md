# Paper 3

**Title:** Quantization-Aware LoRA: 4-bit Adapters That Don't Drift
**Authors:** H. Park, J. Boateng, V. Sørensen, A. Rossi
**Venue:** Workshop on Parameter-Efficient Tuning (PET '25)
**Year:** 2025

## Abstract

Low-rank adapters (LoRA) degrade sharply when the base model is quantized to 4
bits. We introduce QA-LoRA, which trains adapters with the quantization error in
the loop, eliminating the post-quantization accuracy cliff and keeping the model
within 1.5% of full-precision fine-tuning.

## Key findings

- Naïve 4-bit LoRA loses 9.7 accuracy points; QA-LoRA loses 1.5.
- Adapter storage stays at 0.3% of base-model size.
- Throughput is identical to standard 4-bit inference (adapters fuse at load).

## Reported accuracy retention (for charting)

- Full-precision fine-tune: 100%
- QA-LoRA (4-bit): 98.5%
- Naïve 4-bit LoRA: 90.3%
