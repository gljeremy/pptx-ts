// Category-related objects. (`pptx.chart.category`)
import { PySequence } from "../shared.js";
import { pyItem } from "../util.js";

/**
 * A sequence of |category.Category| objects, each representing a category
 * label on the chart. Provides properties for dealing with hierarchical
 * categories.
 */
export class Categories extends PySequence<Category> {
  declare _xChart: any;

  constructor(xChart: any) {
    super();
    this._xChart = xChart;
  }

  at(idx: number): Category {
    const pt = pyItem(this._xChart.catPts, idx);
    return new Category(pt, idx);
  }

  override *[Symbol.iterator](): Generator<Category> {
    const catPts = this._xChart.catPts;
    for (const [idx, pt] of [...catPts].entries()) {
      yield new Category(pt, idx);
    }
  }

  get length(): number {
    return this._xChart.catPtCount;
  }

  /**
   * Return an integer representing the number of hierarchical levels in
   * this category collection. Returns 1 for non-hierarchical categories
   * and 0 if no categories are present (generally meaning no series are
   * present).
   */
  get depth() {
    const cat = this._xChart.cat;
    if (cat === null) {
      return 0;
    }
    if (cat.multiLvlStrRef === null) {
      return 1;
    }
    return cat.lvls.length;
  }

  /**
   * Return a sequence of tuples, each containing the flattened hierarchy
   * of category labels for a leaf category. Each tuple is in parent ->
   * child order, e.g. ``('US', 'CA', 'San Francisco')``, with the leaf
   * category appearing last. If this categories collection is
   * non-hierarchical, each tuple will contain only a leaf category label.
   * If the plot has no series (and therefore no categories), an empty
   * tuple is returned.
   */
  get flattenedLabels(): string[][] {
    const cat = this._xChart.cat;
    if (cat === null) {
      return [];
    }
    if (cat.multiLvlStrRef === null) {
      return [...[...this].map((category) => [category.label])];
    }
    return [...this._iterFlattenedCategories()].map((flatCat) => [...flatCat].reverse().map((category) => category.label));
  }

  /**
   * Return a sequence of |CategoryLevel| objects representing the
   * hierarchy of this category collection. The sequence is empty when the
   * category collection is not hierarchical, that is, contains only
   * leaf-level categories. The levels are ordered from the leaf level to
   * the root level; so the first level will contain the same categories
   * as this category collection.
   */
  get levels(): CategoryLevel[] {
    const cat = this._xChart.cat;
    if (cat === null) {
      return [];
    }
    return [...cat.lvls].map((lvl) => new CategoryLevel(lvl));
  }

  /**
   * Generate a ``tuple`` object for each leaf category in this
   * collection, containing the leaf category followed by its "parent"
   * categories, e.g. ``('San Francisco', 'CA', 'USA'). Each tuple will be
   * the same length as the number of levels (excepting certain edge
   * cases which I believe always indicate a chart construction error).
   */
  *_iterFlattenedCategories(): Generator<Category[]> {
    const levels = this.levels;
    if (levels.length === 0) {
      return;
    }
    const [leafLevel, remainingLevels] = [levels[0], levels.slice(1)];
    for (const category of leafLevel) {
      yield this._parentage([category], remainingLevels);
    }
  }

  /**
   * Return a tuple formed by recursively concatenating *categories* with
   * its next ancestor from *levels*. The idx value of the first category
   * in *categories* determines parentage in all levels. The returned
   * sequence is in child -> parent order. A parent category is the
   * Category object in a next level having the maximum idx value not
   * exceeding that of the leaf category.
   */
  _parentage(categories: Category[], levels: CategoryLevel[]): Category[] {
    let parent;
    if (levels.length === 0) {
      return [...categories];
    }
    if (levels[0].length === 0) {
      return [...categories];
    }
    const [parentLevel, remainingLevels] = [levels[0], levels.slice(1)];
    const leafNode = categories[0];
    parent = parentLevel.at(0);
    for (const category of parentLevel) {
      if (category.idx > leafNode.idx) {
        break;
      }
      parent = category;
    }
    const extendedCategories = [...categories, parent];
    return this._parentage(extendedCategories, remainingLevels);
  }
}

/**
 * An extension of `str` that provides the category label as its string
 * value, and additional attributes representing other aspects of the
 * category.
 *
 * python's `str` subclass is a `String` object here: `String(category)`, template strings and
 * `category == "label"` give the label.
 */
export class Category extends String {
  declare _element: any;
  declare _pt: any;
  declare _idx: any;

  /**
   * *idx* is a required attribute of a c:pt element, but must be
   * specified when pt is None, as when a "placeholder" category is
   * created to represent a missing c:pt element.
   */
  constructor(pt: any, idx: any = null) {
    super(pt === null || pt === undefined ? "" : pt.v.text);
    this._element = pt;
    this._pt = pt;
    this._idx = idx;
  }

  /**
   * Return an integer representing the index reference of this category.
   * For a leaf node, the index identifies the category. For a parent (or
   * other ancestor) category, the index specifies the first leaf category
   * that ancestor encloses.
   */
  get idx() {
    if (this._pt === null) {
      return this._idx;
    }
    return this._pt.idx;
  }

  /** Return the label of this category as a string. */
  get label(): string {
    return this.valueOf();
  }
}

/**
 * A sequence of |category.Category| objects representing a single level in
 * a hierarchical category collection. This object is only used when the
 * categories are hierarchical, meaning they have more than one level and
 * higher level categories group those at lower levels.
 */
export class CategoryLevel extends PySequence<Category> {
  declare _element: any;
  declare _lvl: any;

  constructor(lvl: any) {
    super();
    this._element = lvl;
    this._lvl = lvl;
  }

  at(offset: number): Category {
    return new Category(pyItem(this._lvl.ptLst, offset));
  }

  get length(): number {
    return this._lvl.ptLst.length;
  }
}
