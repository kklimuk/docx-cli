# Paper 1

**Title:** Sparse Mixture-of-Experts Routing for Edge Inference
**Authors:** L. Nguyen, R. Castellano, P. Adeyemi
**Venue:** Proceedings of the 2025 Conference on Efficient Machine Learning (EML '25)
**Year:** 2025

## Abstract

We present GateLite, a sparse mixture-of-experts (MoE) router designed for
inference on memory-constrained edge devices. By activating at most two of eight
expert subnetworks per token, GateLite matches the accuracy of a dense baseline
while reducing peak memory by 61%.

## Key findings

- GateLite reaches 94.2% of dense-model accuracy on the GLUE-Edge suite.
- Peak memory drops from 3.1 GB (dense) to 1.2 GB (GateLite).
- Median latency improves 2.4× on a Cortex-A78 reference board.

## Reported accuracy retention (for charting)

- Dense baseline: 100%
- GateLite (2-of-8): 94.2%
- Static prune (50%): 81.0%
