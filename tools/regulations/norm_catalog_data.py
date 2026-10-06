"""Strict, invariant norm matrices transcribed from VBHN97-2025-BQP Appendix I."""

from __future__ import annotations


def _data(variants: str, rates: str, adjustments: str = "", constraints: str = "") -> str:
    fields = [f"variantCodes={variants}", f"rates={rates}"]
    if adjustments:
        fields.append(f"adjustments={adjustments}")
    if constraints:
        fields.append(f"constraints={constraints}")
    return ";".join(fields)


SLOPE = "factor:slope-gt-25deg:Labor:*:ratio:1.10"
CURRENT = "|".join([
    "factor:current-gt-0-le-0.5:Labor+Machine:*:ratio:1.10",
    "factor:current-gt-0.5-le-1:Labor+Machine:*:ratio:1.25",
    "factor:current-gt-1-le-2:Labor+Machine:*:ratio:1.50",
])


NORM_DATA = {
    "NORM-000.0100": _data(
        "plain-midland,mountain-island",
        "Labor:LAB-TEAM-LEADER:worker-day:2,3|"
        "Labor:LAB-INVESTIGATOR:worker-day:15,22.5|"
        "Labor:LAB-SUPPORT:worker-day:2,3"),
    "NORM-000.0200": _data(
        "forest-1,forest-2,forest-3,forest-4",
        "Labor:LAB-QNCN-7:worker-day:33.5,37.5,41.5,50.5"),
    "NORM-000.0300": _data(
        "standard",
        "Material:MAT-WOOD-STAKE:each:0.5|Material:MAT-ROPE-10MM:m:67|"
        "Material:MAT-RED-FLAG:each:2|Material:MAT-OTHER:percent:5|"
        "Labor:LAB-QNCN-8:worker-day:2.31|Machine:M010.001:shift:9.25|"
        "Machine:M010.006:shift:1"),
    "NORM-000.0400": _data(
        "soil-1,soil-2,soil-3,soil-4",
        "Labor:LAB-QNCN-8:worker-day:0.04,0.044,0.049,0.053|"
        "Machine:M010.001:shift:0.014,0.014,0.014,0.014"),
    "NORM-010.0100": _data(
        "forest-1,forest-2,forest-3,forest-4",
        "Labor:LAB-QNCN-8:worker-day:111,125,143,167", SLOPE),
    "NORM-010.0200": _data(
        "forest-1,forest-2,forest-3,forest-4",
        "Labor:LAB-QNCN-7:worker-day:67,73,83,101", SLOPE),
    "NORM-010.0300": _data(
        "forest-1,forest-2,forest-3,forest-4",
        "Material:MAT-GASOLINE-OR-DIESEL:kg:420,448,476,546|"
        "Labor:LAB-QNCN-7:worker-day:67,73,83,101", SLOPE),
    "NORM-010.0400": _data(
        "single-fence,double-fence",
        "Material:MAT-TNT:kg:2,3|Material:MAT-SLOW-FUSE:m:0.5,1|"
        "Material:MAT-ELECTRIC-DETONATOR:each:1.2,1.2|"
        "Material:MAT-ELECTRIC-WIRE:m:2,2|Material:MAT-WOOD-STRIP:m:1,1|"
        "Material:MAT-BARBED-WIRE:kg:0.1,0.15|Material:MAT-SIGN:each:0.08,0.08|"
        "Material:MAT-OTHER:percent:5,5|Labor:LAB-QNCN-8:worker-day:0.17,0.25|"
        "Machine:M010.025:shift:0.014,0.014|Machine:M010.026:shift:0.014,0.014",
        SLOPE),
    "NORM-020.0100": _data(
        "density-1,density-2,density-3,density-4",
        "Material:MAT-CONCRETE-STAKE:each:4,4,4,4|Material:MAT-WOOD-STAKE:each:2,2,2,2|"
        "Material:MAT-WHITE-FLAG:each:90,90,90,90|Material:MAT-RED-FLAG:each:2,4,6,8|"
        "Material:MAT-SIGN:each:0.4,0.4,0.4,0.4|Material:MAT-OTHER:percent:5,5,5,5|"
        "Labor:LAB-QNCN-8:worker-day:167,182,200,223"),
    "NORM-020.0200": _data(
        "density-1,density-2,density-3,density-4",
        "Material:MAT-CONCRETE-STAKE:each:4,4,4,4|Material:MAT-WOOD-STAKE:each:34,34,34,34|"
        "Material:MAT-ROPE-10MM:m:67,67,67,67|Material:MAT-RED-FLAG:each:2,4,6,8|"
        "Material:MAT-OTHER:percent:5,5,5,5|Labor:LAB-QNCN-7:worker-day:17.36,19.10,21,23.10|"
        "Machine:M010.001:shift:11.57,12.73,14,15.4"),
    "NORM-020.0300": _data(
        "soil-1,soil-2,soil-3,soil-4",
        "Labor:LAB-QNCN-8:worker-day:0.060,0.078,0.107,0.160|"
        "Machine:M010.001:shift:0.014,0.014,0.014,0.014",
        "add:uxo-signal:Labor:LAB-QNCN-8:worker-day:0.028"),
    "NORM-020.0400": _data(
        "soil-1,soil-2,soil-3,soil-4",
        "Labor:LAB-QNCN-8:worker-day:0.070,0.094,0.132,0.199|"
        "Machine:M010.001:shift:0.014,0.014,0.014,0.014",
        "add:uxo-signal:Labor:LAB-QNCN-8:worker-day:0.028"),
    "NORM-020.0500": _data(
        "depth-0.5-or-1,depth-3,depth-5,depth-10",
        "Material:MAT-WOOD-STAKE:each:50,50,50,50|Material:MAT-ROPE-10MM:m:100,100,100,100|"
        "Material:MAT-RED-FLAG-LARGE:each:1,1,1,1|Material:MAT-OTHER:percent:5,5,5,5|"
        "Labor:LAB-QNCN-7:worker-day:6.40,7.05,7.76,8.54|"
        "Machine:M010.002:shift:4.27,4.70,5.17,0|Machine:M010.003:shift:0,0,0,5.69",
        "", "requires-condition:depth-0.5-or-1:forestry-salt-independent-or-owner-request"),
    "NORM-020.0600": _data(
        "soil-1,soil-2,soil-3,soil-4",
        "Labor:LAB-QNCN-8:worker-day:0.71,1.04,1.51,2.34|"
        "Machine:M010.002:shift:0.008,0.008,0.008,0.008",
        "add:water-excavation:Machine:M010.023:shift:0.012"),
    "NORM-020.0700": _data(
        "soil-1,soil-2,soil-3,soil-4",
        "Material:MAT-BAMBOO-STAKE:each:0.2,0.2,0,0|Material:MAT-WOOD-BOARD:m3:0.004,0.004,0,0|"
        "Material:MAT-NAIL-10CM:kg:0.15,0.15,0,0|Material:MAT-OTHER:percent:1,1,1,1|"
        "Labor:LAB-QNCN-8:worker-day:0.78,1.14,1.66,2.57|"
        "Machine:M010.002:shift:0.008,0.008,0.008,0.008",
        "add:water-excavation:Machine:M010.023:shift:0.012"),
    "NORM-020.0800": _data(
        "soil-1,soil-2,soil-3,soil-4",
        "Material:MAT-BAMBOO-STAKE:each:0.4,0.4,0.2,0.1|Material:MAT-WOOD-BOARD:m3:0.008,0.008,0.004,0.002|"
        "Material:MAT-NAIL-TWO-PRONG:kg:0.4,0.4,0.2,0|Material:MAT-NAIL-10CM:kg:0.15,0.15,0.10,0.05|"
        "Material:MAT-WIRE-2MM:kg:0.20,0.15,0.12,0.01|Material:MAT-OTHER:percent:1,1,1,1|"
        "Labor:LAB-QNCN-8:worker-day:0.66,0.88,1.28,1.87|"
        "Machine:M010.002:shift:0.008,0.008,0.008,0.008",
        "add:water-excavation:Machine:M010.023:shift:0.012"),
    "NORM-020.0900": _data(
        "soil-1,soil-2,soil-3,soil-4",
        "Material:MAT-BAMBOO-STAKE:each:0.2,0.2,0.1,0|Material:MAT-WOOD-BOARD:m3:0.004,0.004,0.002,0|"
        "Material:MAT-NAIL-TWO-PRONG:kg:0.4,0.4,0.2,0|Material:MAT-NAIL-10CM:kg:0.15,0.15,0.10,0|"
        "Material:MAT-WIRE-2MM:kg:0.20,0.15,0.12,0|Material:MAT-OTHER:percent:1,1,1,0|"
        "Labor:LAB-QNCN-7:worker-day:0.0300,0.0390,0.0461,0.0518|"
        "Labor:LAB-QNCN-8:worker-day:0.72,0.96,1.40,2.04|"
        "Machine:M010.002-OR-M010.003:shift:0.008,0.008,0.008,0.008|"
        "Machine:M010.004:shift:0.0053,0.0063,0.0090,0.0104",
        "add:water-excavation:Machine:M010.023:shift:0.012"),
    "NORM-020.1000": _data(
        "standard",
        "Material:MAT-SMOOTH-PLASTIC-TUBE:m:125|Material:MAT-WOOD-STAKE:each:50|"
        "Material:MAT-ROPE-10MM:m:100|Material:MAT-RED-FLAG-LARGE:each:1|"
        "Material:MAT-OTHER:percent:1|Labor:LAB-QNCN-7:worker-day:7.76|"
        "Machine:M010.002:shift:5.17",
        "", "external-estimate:drilling-cost"),
    "NORM-020.1100": _data(
        "soil-1,soil-2,soil-3,soil-4",
        "Material:MAT-BAMBOO-STAKE:each:0.2,0.2,0.1,0|Material:MAT-WOOD-BOARD:m3:0.004,0.004,0.002,0|"
        "Material:MAT-NAIL-TWO-PRONG:kg:0.4,0.4,0.2,0|Material:MAT-NAIL-10CM:kg:0.15,0.15,0.10,0|"
        "Material:MAT-WIRE-2MM:kg:0.20,0.15,0.12,0|Material:MAT-OTHER:percent:1,1,1,0|"
        "Labor:LAB-QNCN-7:worker-day:0.0300,0.0390,0.0461,0.0518|"
        "Labor:LAB-QNCN-8:worker-day:0.72,0.96,1.40,2.04|"
        "Machine:M010.002-OR-M010.003:shift:0.008,0.008,0.008,0.008|"
        "Machine:M010.004:shift:0.0053,0.0063,0.0090,0.0104",
        "add:water-excavation:Machine:M010.023:shift:0.012"),
    "NORM-020.1200": _data(
        "weight-lt-3,weight-3-15,weight-15-50,weight-50-120,weight-120-250,weight-gt-250",
        "Material:MAT-TNT:kg:0.2,0.4,1,4,4,4|Material:MAT-ELECTRIC-DETONATOR:each:1,1,1,1,1,1|"
        "Material:MAT-ELECTRIC-WIRE:m:2,4,6,10,20,20|Material:MAT-CLOTH:m2:0,0,0.4,1,1,1|"
        "Material:MAT-TWINE:kg:0,0,0.1,0.1,0.1,0.1|Material:MAT-SIGN:each:0.08,0.08,0.08,0.08,0.08,0.08|"
        "Material:MAT-OTHER:percent:1,1,1,1,1,1|Labor:LAB-QNCN-7:worker-day:0.12,0.12,0.14,0.14,0.17,0.17|"
        "Labor:LAB-QNCN-8:worker-day:0.060,0.060,0.070,0.070,0.085,0.085|"
        "Machine:M010.025:shift:0.014,0.014,0.014,0.014,0.014,0.014|"
        "Machine:M010.026:shift:0.014,0.014,0.014,0.014,0.014,0.014"),
    "NORM-030.0100": _data(
        "water-0.5-12,water-12-22,water-22-25",
        "Material:MAT-ANCHOR-50KG:each:0.032,0.032,0.032|Material:MAT-ANCHOR-20KG:each:0.064,0.064,0.064|"
        "Material:MAT-LARGE-FLOAT:each:0.32,0.32,0.32|Material:MAT-SMALL-FLOAT:each:4,4,4|"
        "Material:MAT-ROPE-10MM:m:210,210,210|Material:MAT-ROPE-12MM:m:8,18.5,27|"
        "Material:MAT-ROPE-18MM:m:8.64,8.64,8.64|Material:MAT-OTHER:percent:1,1,1|"
        "Labor:LAB-QNCN-7:worker-day:23.82,26.19,28.8|Machine:M010.008:shift:9.26,10.19,12.22|"
        "Machine:M010.010:shift:9.26,10.19,12.22|Machine:M010.009:shift:18.52,20.38,24.44|"
        "Machine:M010.027:shift:10.19,12.12,13.44", CURRENT, "prohibit-condition:current-gt-2"),
    "NORM-030.0200": _data(
        "water-0.5-12,water-12-22,water-22-25",
        "Material:MAT-ANCHOR-50KG:each:0.032,0.032,0.032|Material:MAT-ANCHOR-20KG:each:0.064,0.064,0.064|"
        "Material:MAT-LARGE-FLOAT:each:0.32,0.32,0.32|Material:MAT-SMALL-FLOAT:each:4,4,4|"
        "Material:MAT-ROPE-10MM:m:105,105,105|Material:MAT-ROPE-12MM:m:8,18.5,27|"
        "Material:MAT-ROPE-18MM:m:8.64,8.64,8.64|Material:MAT-OTHER:percent:1,1,1|"
        "Labor:LAB-QNCN-7:worker-day:9.81,10.8,11.88|Machine:M010.008:shift:4.63,5.10,5.61|"
        "Machine:M010.010:shift:4.63,5.10,5.61|Machine:M010.009:shift:9.26,10.20,11.22|"
        "Machine:M010.027:shift:5.10,5.61,6.17", CURRENT, "prohibit-condition:current-gt-2"),
    "NORM-030.0300": _data(
        "water-0.5-3,water-3-12,water-12-22,water-22-25",
        "Material:MAT-SPECIAL-ANCHOR:each:0,0.002,0.002,0.002|Material:MAT-PLASTIC-FLOAT:each:0,0.067,0.067,0.067|"
        "Material:MAT-ROPE-12MM:m:0,0.16,0.38,0.56|Material:MAT-BAMBOO-POLE:each:0.2,0,0,0|"
        "Material:MAT-OTHER:percent:1,1,1,1|Labor:LAB-QNCN-7:worker-day:0.056,0.056,0.067,0.081|"
        "Machine:M010.008:shift:0.014,0.016,0.017,0.019|Machine:M010.010:shift:0.014,0.016,0.017,0.019",
        CURRENT, "prohibit-condition:current-gt-2"),
    "NORM-030.0400": _data(
        "water-0.5-12,water-12-22,water-22-25",
        "Labor:LAB-QNCN-7:worker-day:0.23,0.26,0.29|Machine:M010.008:shift:0.014,0.015,0.017|"
        "Machine:M010.DIVING:shift:0.193,0.212,0.233|Machine:M010.027:shift:0.207,0.227,0.250",
        CURRENT, "prohibit-condition:current-gt-2"),
    "NORM-030.0500": _data(
        "water-0.5-12,water-12-22,water-22-25",
        "Labor:LAB-QNCN-7:worker-day:0.23,0.26,0.29|Machine:M010.008:shift:0.014,0.015,0.017|"
        "Machine:M010.010:shift:0.014,0.015,0.017|Machine:M010.DIVING:shift:0.153,0.168,0.185|"
        "Machine:M010.023:shift:0.125,0.138,0.152|Machine:M010.027:shift:0.198,0.213,0.230",
        CURRENT, "prohibit-condition:current-gt-2"),
    "NORM-030.0600": _data(
        "water-0.5-12,water-12-22,water-22-25",
        "Material:MAT-STEEL-EXCAVATION-FRAME:kg:24.32,24.32,24.32|Labor:LAB-QNCN-7:worker-day:1,1.1,1.21|"
        "Machine:M010.008:shift:0.014,0.015,0.017|Machine:M010.010:shift:0.014,0.015,0.017|"
        "Machine:M010.DIVING:shift:0.264,0.290,0.319|Machine:M010.023:shift:0.125,0.138,0.152|"
        "Machine:M010.027:shift:0.419,0.445,0.474", CURRENT, "prohibit-condition:current-gt-2"),
    "NORM-030.0700": _data(
        "water-0.5-12,water-12-22,water-22-25",
        "Material:MAT-STEEL-EXCAVATION-FRAME:kg:40.5,40.5,40.5|Labor:LAB-QNCN-7:worker-day:1.33,1.46,1.61|"
        "Machine:M010.008:shift:0.014,0.015,0.017|Machine:M010.010:shift:0.014,0.015,0.017|"
        "Machine:M010.DIVING:shift:0.640,0.704,0.774|Machine:M010.023:shift:0.420,0.462,0.510|"
        "Machine:M010.027:shift:0.640,0.704,0.774", CURRENT, "prohibit-condition:current-gt-2"),
    "NORM-030.0800": _data(
        "weight-lt-3,weight-3-15,weight-15-50,weight-50-120,weight-120-250,weight-gt-250",
        "Material:MAT-TNT:kg:0.2,0.4,1,4,4,4|Material:MAT-ELECTRIC-DETONATOR:each:1,1,2,2,2,2|"
        "Material:MAT-ELECTRIC-WIRE:m:2,4,6,10,20,30|Material:MAT-CLOTH:m2:0,0,0.4,1,1,1|"
        "Material:MAT-TWINE:kg:0,0.1,0.15,0.2,0.2,0.2|Material:MAT-SIGN:each:0.08,0.08,0.08,0.08,0.08,0.08|"
        "Material:MAT-OTHER:percent:1,1,1,1,1,1|Labor:LAB-QNCN-7:worker-day:2.04,2.04,2.04,2.34,2.34,2.34|"
        "Labor:LAB-QNCN-8:worker-day:0.34,0.34,0.34,0.39,0.39,0.39|"
        "Machine:M010.008:shift:0.014,0.014,0.014,0.014,0.014,0.014|"
        "Machine:M010.025:shift:0.014,0.014,0.014,0.014,0.014,0.014|"
        "Machine:M010.026:shift:0.014,0.014,0.014,0.014,0.014,0.014|"
        "Machine:M010.027:shift:0.17,0.17,0.17,0.19,0.19,0.19|"
        "Machine:M010.009:shift:0.34,0.34,0.34,0.39,0.39,0.39|"
        "Machine:M010.DIVING:shift:0.08,0.08,0.08,0.1,0.1,0.1",
        CURRENT, "prohibit-condition:current-gt-2"),
    "NORM-040.0100": _data(
        "standard",
        "Labor:LAB-QNCN-7:worker-day:0.125|Machine:M010.020:shift:0.042|"
        "Machine:M010.021:shift:0.042|Machine:M010.014:shift:0.042|Machine:M010.015:shift:0.042"),
    "NORM-040.0200": _data(
        "standard",
        "Labor:LAB-QNCN-7:worker-day:0.33|Labor:LAB-QNCN-8:worker-day:0.17|"
        "Machine:M010.013:shift:0.085|Machine:M010.015:shift:0.085|Machine:M010.022:shift:0.009|"
        "Machine:M010.020:shift:0.085|Machine:M010.019:shift:0.085|Machine:M010.016:shift:0.085|"
        "Machine:M010.017:shift:0.085"),
    "NORM-040.0300": _data(
        "water-25-50,water-50-150",
        "Labor:LAB-QNCN-7:worker-day:0.92,1.2|Machine:M010.020:shift:0.063,0.063|"
        "Machine:M010.018-OR-M010.028:shift:0.23,0.30|Machine:M010.007:shift:0.14,0|"
        "Machine:M010.010:shift:0.14,0|Machine:M010.014:shift:0.34,0.41|"
        "Machine:M010.015:shift:0.34,0.41|Machine:M010.022:shift:0.035,0.04"),
    "NORM-040.0400": _data(
        "water-25-50",
        "Material:MAT-STEEL-EXCAVATION-FRAME:kg:24.32|Labor:LAB-QNCN-7:worker-day:3.8|"
        "Machine:M010.007:shift:0.14|Machine:M010.023:shift:0.23|Machine:M010.DIVING:shift:0.74|"
        "Machine:M010.024:shift:0.74|Machine:M010.010:shift:0.14|Machine:M010.014:shift:0.74|"
        "Machine:M010.015:shift:0.74|Machine:M010.022:shift:0.07",
        "", "external-estimate:water-depth-gt-50-to-150"),
    "NORM-040.0500": _data(
        "standard",
        "Material:MAT-ANCHOR-CONCRETE-20KG:each:1|Material:MAT-NYLON-ROPE-14MM:m:30|"
        "Material:MAT-FLOAT-40L:each:2|Material:MAT-TNT:kg:6|"
        "Material:MAT-WATERPROOF-DETONATING-CORD:m:66|Material:MAT-ELECTRIC-DETONATOR:each:2|"
        "Material:MAT-ELECTRIC-WIRE:m:20|Material:MAT-CLOTH:m2:1|Material:MAT-TWINE:kg:0.2|"
        "Material:MAT-OTHER:percent:10|Labor:LAB-QNCN-7:worker-day:0.8|Labor:LAB-QNCN-8:worker-day:0.1|"
        "Machine:M010.020:shift:0.063|Machine:M010.DIVING:shift:0.19|Machine:M010.007:shift:0.1|"
        "Machine:M010.010:shift:0.1|Machine:M010.025:shift:0.02|Machine:M010.026:shift:0.02|"
        "Machine:M010.014:shift:0.33|Machine:M010.015:shift:0.33|Machine:M010.022:shift:0.04"),
    "NORM-040.0600": _data(
        "water-25-50,water-50-100,water-100-150",
        "Material:MAT-ANCHOR-50KG:each:1,1,1|Material:MAT-NYLON-ROPE-14MM:m:135,280,425|"
        "Material:MAT-FLOAT-40L:each:4,4,4|Material:MAT-TNT:kg:10,10,10|Material:MAT-CLOTH:m2:1,1,1|"
        "Material:MAT-ELECTRIC-DETONATOR:each:2,2,2|Material:MAT-WATERPROOF-DETONATING-CORD:m:165,330,495|"
        "Material:MAT-TWINE:kg:0.2,0.2,0.2|Material:MAT-ELECTRIC-WIRE:m:20,20,20|"
        "Material:MAT-OTHER:percent:10,10,10|Labor:LAB-QNCN-7:worker-day:0.6,0.76,0.96|"
        "Labor:LAB-QNCN-8:worker-day:0.1,0.1,0.1|Machine:M010.020:shift:0.063,0.063,0.063|"
        "Machine:M010.018-OR-M010.028:shift:0.15,0.19,0.24|Machine:M010.025:shift:0.02,0.02,0.02|"
        "Machine:M010.026:shift:0.02,0.02,0.02|Machine:M010.007:shift:0.08,0.1,0.12|"
        "Machine:M010.010:shift:0.08,0.1,0.12|Machine:M010.014:shift:0.36,0.39,0.45|"
        "Machine:M010.015:shift:0.04,0.04,0.05|Machine:M010.022:shift:0.04,0.04,0.05"),
}
