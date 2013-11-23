module Gini

// Expert F# 3.0, chapter 6, p. 142
module List =
    let rec pairwise l =
        match l with
        | [] | [_] -> []
        | h1 :: ((h2 :: _) as t) -> (h1, h2) :: pairwise t


open Deedle

(**   cum-fn = clojure.core.reductions
(defn cum-fn
  ([s cfn]
     (cond (empty? s) nil
           (empty? (rest s)) (list (first s))
           :else (lazy-seq (cons (first s) (cum-fn (first s) (rest s) cfn)))))
  ([x s cfn]
     (cond (empty? s) nil
           (empty? (rest s)) (list (cfn x (first s)))
           :else (lazy-seq (cons (cfn x (first s)) (cum-fn (cfn x (first s)) (rest s) cfn))))))
**)
let reductions (s:List<'a>) (fn: 'a -> 'b) : List<'b> =
    let rec reduc1 (x:'b) (s:List<'a>) (fn: 'a -> 'b) : List<'b> = 
        if s.IsEmpty then []
        elif s.Tail.IsEmpty then [(fn s.Head)]
        else [(fn s.Head)] @ (reduc1 (fn s.Head) s.Tail fn)
    if s.IsEmpty then []
    elif s.Tail.IsEmpty then [(fn s.Head)]
    else [(fn s.Head)] @ (reduc1 (fn s.Head) s.Tail fn)



let calcGini (frame: Frame<_,_>) selector (missingValue: float) : float =
    let series = 
        frame.GetSeries selector 
        |> Series.fillMissingWith missingValue  
        |> Series.map (fun x y -> float y) 
    let series_sorted = series.Values |> Seq.sort
    let max = series_sorted |> Seq.last
    let z = 0.0 :: reductions (List.ofSeq series_sorted) (fun x -> x / max) 
    let y = 0.0 :: reductions [1..z.Length-1] (fun x -> (float x) / (float (z.Length-1)))
    let x = List.zip y z |> List.pairwise |> List.map (fun ((x0,y0), (x1,y1)) -> (x1 - x0) * (y1 + y0))
    1.0 - List.reduce (+) x 