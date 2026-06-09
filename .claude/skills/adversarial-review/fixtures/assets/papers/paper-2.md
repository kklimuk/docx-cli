# Paper 2

**Title:** Retrieval-Conditioned Distillation for Small Language Models
**Authors:** S. Okonkwo, M. Lindqvist
**Venue:** Transactions on Compact Models, Vol. 4
**Year:** 2024

## Abstract

We distill a 70B-parameter teacher into a 1.3B student by conditioning the
distillation loss on retrieved context. The retrieval-conditioned student (RCD)
closes 78% of the quality gap to the teacher on open-domain QA while running on a
single consumer GPU.

## Key findings

- RCD recovers 78% of the teacher–student quality gap versus 41% for vanilla
  distillation.
- Adding retrieval at train time, not just test time, accounts for two-thirds of
  the gain.
- The 1.3B student answers 19% more open-domain QA items correctly than a
  same-size model trained from scratch.

## Reported accuracy retention (for charting)

- Teacher (70B): 100%
- RCD student (1.3B): 78%
- Vanilla distillation (1.3B): 41%
