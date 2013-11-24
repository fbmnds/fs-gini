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


//let reductions (s:List<'a>) (fn: 'a -> 'b) : List<'b> =
//    let rec reduc1 (x:'b) (s:List<'a>) (fn: 'a -> 'b) : List<'b> =
//        if s.IsEmpty then []
//        elif s.Tail.IsEmpty then [(fn s.Head)]
//        else [(fn s.Head)] @ (reduc1 (fn s.Head) s.Tail fn)
//    if s.IsEmpty then []
//    elif s.Tail.IsEmpty then [(fn s.Head)]
//    else [(fn s.Head)] @ (reduc1 (fn s.Head) s.Tail fn)

let reductions (fn:'b -> 'a -> 'b) (init:'b) (s:'a list) : 'b list =
    let rec reduc (fn:'b -> 'a -> 'b) (x:'b list) (xLast: 'b) (s:'a list) : 'b list =
        if s.Tail.IsEmpty then x @ [(fn xLast s.Head)]
        else reduc fn (x @ [(fn xLast s.Head)]) (fn xLast s.Head) s.Tail
    if s.IsEmpty then [init]
    elif s.Tail.IsEmpty then [(fn init s.Head)]
    else (reduc fn [(fn init s.Head)] (fn init s.Head) s.Tail)


let vecAdd (x:float*float) (y:float*float) = ((fst x)+(fst y),(snd x)+(snd y))

let calcGini (frame: Frame<_,_>) selector (missingValue: float) : float =

    let series = 
        frame.GetSeries selector 
        |> Series.fillMissingWith missingValue  
        |> Series.map (fun x y -> float y) 

    let normValue = series.Values |> Seq.sum
    let sortedValues = series.Values |> Seq.sort |> List.ofSeq |> List.map (fun x -> x / normValue)
    
    let len = sortedValues.Length
    let sortedXY = 
        List.zip (0.0 :: (List.init len (fun i -> 1.0 / float len))) (0.0 :: sortedValues)
        |> reductions vecAdd (0.0,0.0)
        |> List.pairwise 
        |> List.map (fun ((x0,y0), (x1,y1)) -> (x1 - x0) * (y1 + y0))

    1.0 - List.reduce (+) sortedXY