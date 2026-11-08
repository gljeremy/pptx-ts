// Stand-ins for `pptx.chart` classes, which are ported in stage 5 (gljeremy/pptx-ts#5). Tests
// before then only use them as specs for mocks (`instance_mock(request, ChartData)`); the
// fixups in test/fixups.json point those imports here. Remove when stage 5 lands.

/** `pptx.chart.data.ChartData` (mock spec only). */
export class ChartData {
  xmlBytes(_chartType: any): Uint8Array {
    return new Uint8Array();
  }
  get xlsxBlob(): Uint8Array {
    return new Uint8Array();
  }
}

/** `pptx.chart.chart.Chart` (mock spec only). */
export class Chart {
  get chartType(): any {
    return null;
  }
}
