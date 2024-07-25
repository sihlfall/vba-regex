# Avoiding infinite loops in the DFS engine

## Problem

If the atom underlying a Kleene star expression matches the empty string, a DFS regex engine
may run into an infinite loop.

## Goal

An atom matching the empty string should be matched against only once in the greedy case (`*`).
In the humble case (`*?`), it should not match at all.

## How can we achieve this?

We can remember the current string pointer (`sp`) value when entering the atom. On leaving the atom, we can then
compare the current string pointer value to the remembered string pointer value. If the string pointer has not
been advanced compared to the remembered value, the atom has matched the empty string. In this case,
we continue processing with the instruction after the Kleene star expression (thus in fact transforming
the Kleene star into question mark expression having matched successfully).

A straight forward implementation of remembering string pointer values requires a stack, since Kleene star
expressions can be nested.

Can we avoid building a stack?

The only reason for the process counter `pc` to decrease while
following a path in the search tree is a `REPEAT_END` instruction at the end of a Kleene star 
or finite quantifier iteration. Furthermore, we never jump from inside a loop to the outside.

We can, therefore, do the following:

* We continuously remember the latest process counter at which the string pointer was advanced (variable `pcadv`,
  which is set to `-1` initially).
* When we reach a `REPEAT_END`, we compare `pcadv` to the `REPEAT_END` target `t` (the location of the 
  corresponding `SPLIT` or `REPEAT_BEGIN` instruction). As we show below, the atom
  has matched the empty string if and only if `pcadv <= t`. In this case, we do not jump backwards to the beginning of
  the loop, yet
  rather continue with the instruction following the loop immediately.
* In the other case, _after_ jumping backwards, however, `pcadv` would subsequently be larger than the process counter.
  In this case, we set `pcadv` to `t`, imagining the target `SPLIT` or `REPEAT_BEGIN` instruction
  (representing the entire previous loop iteration) to have advanced the string pointer.

It remains to be shown that

_Iff `pcadv <= t`, the atom has matched the empty string._

Proof. Suppose `pcadv > t`. Since between `t` and `pc`, there can be no instruction jumping over `pc` (as we 
do not jump from the inside of a loop to the outside). We therefore must have `t < pcadv <= pc`.
Since at the start of the current iteration, `pcadv` was smaller than or equal to `t`, the string pointer must 
have been advanced during the current
loop iteration, which means the atom has not matched the empty string.

Suppose `pcadv <= t`. Since loops are nested, `pcadv` cannot have gained its value during the current iteration
of the loop. Hence the string pointer was not advanced during the current iteration, and thus the atom must have
matched the empty string.

## The humble case

A humble Kleene star or humble max finite quantifier iteration matching the empty string indicates a failure path,
since we have tried to find a
match not considering the atom at all first. Only if this has failed we will reach a path in which we try to match
the atom. Yet after successfully matching the empty string to the atom, we will be on the failing path again.
This means we can break immediately after matching the empty string.


