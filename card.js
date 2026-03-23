/* =====================================================
   AgriCard Stine – Card Generator v8.1
   Renderização pixel-perfect baseada em PPTX

   ESTRUTURA DO PPTX:
   - slideMaster1.xml : imagem de fundo ({{layout}})
   - slideLayoutN.xml : posições/fontes de todos os placeholders
                        + placeholder {{logo_variedade}} (type=pic)
                        + imagem do logo via rId (image6.png no v2)
   - slide1.xml       : apenas referências (placeholders vazios)

   PLACEHOLDERS DE TEXTO:
   {{data_plantio}}      {{data_colheita}}
   {{produtividade_int}} {{produtividade_dec}}
   {{unidade}}           {{cidade}}/{{estado}}
   {{produtor}}          {{area}}

   PLACEHOLDER DE IMAGEM (renderizado como img no canvas):
   {{logo_variedade}}  — SP type=pic no slideLayout

   SEPARAÇÃO AUTOMÁTICA DA PRODUTIVIDADE:
   "187,1" → produtividade_int="187" | produtividade_dec=",1"
   "187"   → produtividade_int="187" | produtividade_dec=""
   ===================================================== */

'use strict';

/* ─────────────────────────────────────────────────────────────
   PPTX PARSER v8.1
   Lê o PPTX e extrai:
   1. Imagem de fundo (slideMaster → maior imagem no rels)
   2. Logo da variedade (slideLayout → SP type=pic → imagem via rId)
   3. Elementos de texto (slideLayout → SP com {{placeholder}} no nome)
   4. Dimensões EMU do slide
───────────────────────────────────────────────────────────── */
const PptxParser = {

  /**
   * Parseia um arquivo PPTX e retorna todos os dados necessários para gerar cards.
   *
   * ESTRATÉGIA v9 (preservação total do layout):
   * - template_image = slide COMPLETO renderizado como PNG (fundo + logos + ícones)
   * - logoData = null (logo já está embutida no template_image)
   * - elements = apenas placeholders de TEXTO para substituição dinâmica
   *
   * @param {File|ArrayBuffer} fileOrBuffer
   * @returns {{ slideW, slideH, bgImageData, logoData, thumbnailData, elements[], fullSlideImage }}
   */
  async parseFile(fileOrBuffer) {
    if (!window.JSZip) {
      await this._loadScript('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js');
    }

    const ab  = fileOrBuffer instanceof ArrayBuffer
      ? fileOrBuffer
      : await fileOrBuffer.arrayBuffer();
    const zip = await window.JSZip.loadAsync(ab);

    // ── 1. Dimensões do slide ──────────────────────────────
    const presXml = await zip.file('ppt/presentation.xml')?.async('text') || '';
    const sldSzM  = presXml.match(/<p:sldSz[^>]+cx="(\d+)"[^>]+cy="(\d+)"/);
    const slideW  = sldSzM ? parseInt(sldSzM[1]) : 6858000;
    const slideH  = sldSzM ? parseInt(sldSzM[2]) : 12193588;

    // ── 2. Descobre qual slideLayout o slide usa ───────────
    const slide1RelsTxt  = await zip.file('ppt/slides/_rels/slide1.xml.rels')?.async('text') || '';
    const layoutFile     = this._findLayoutFile(slide1RelsTxt);
    const layoutXml      = await zip.file(layoutFile)?.async('text') || '';
    const layoutRelsFile = layoutFile.replace('slideLayouts/', 'slideLayouts/_rels/').replace('.xml', '.xml.rels');
    const layoutRels     = await zip.file(layoutRelsFile)?.async('text') || '';

    // ── 3. XMLs auxiliares ────────────────────────────────
    const masterRels = await zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels')?.async('text') || '';
    const masterXml  = await zip.file('ppt/slideMasters/slideMaster1.xml')?.async('text') || '';
    const slide1Xml  = await zip.file('ppt/slides/slide1.xml')?.async('text') || '';
    const slide1Rels = slide1RelsTxt;

    // ── 4. Imagem de fundo (maior imagem do master) ───────
    const bgImageData = await this._extractBgFromMaster(zip, masterRels);

    // ── 5. Coleta TODAS as imagens fixas com posições EMU ─
    // Estas imagens (logos, ícones) serão compostas diretamente no template_image.
    // O CardRenderer NÃO reinjetará imagens — apenas sobreporá texto.
    const masterImages = await this._extractImagesWithPositions(zip, masterXml, masterRels);
    const layoutImages = await this._extractImagesWithPositions(zip, layoutXml, layoutRels);
    const slide1Images = await this._extractImagesWithPositions(zip, slide1Xml, slide1Rels);

    // Descobre o fname da imagem usada como fundo (para não desenhá-la duas vezes)
    // _extractBgFromMaster retorna a maior imagem do master — identificamos pelo tamanho
    // A imagem do master que tem cx ≈ 100% da largura é o fundo
    const bgFnames = new Set(masterImages
      .filter(img => img.cx >= slideW * 0.90) // >= 90% da largura = fundo
      .map(img => img.fname));

    // Deduplica pelo fname e exclui o fundo (já desenhado separadamente)
    const seen = new Set([...bgFnames]); // pré-popula com o fundo para excluir
    const allImages = [];
    for (const img of [...masterImages, ...layoutImages, ...slide1Images]) {
      if (seen.has(img.fname)) continue;
      seen.add(img.fname);
      allImages.push(img);
    }

    console.log(`[PptxParser] Imagens fixas: ${allImages.map(i => i.fname.replace('ppt/media/', '')).join(', ') || 'nenhuma'}`);

    // ── 6. Renderiza o slide COMPLETO como PNG ─────────────
    // fundo + logos + ícones — preserva transparência (PNG, não JPEG)
    const fullSlideImage = await this.renderFullSlide(zip, slideW, slideH, bgImageData, allImages, 1080);

    // ── 7. Mapa de cores do tema (resolve schemeClr no canvas) ─
    const themeColors = await this._extractThemeColors(zip);

    // ── 8. Placeholders de texto ──────────────────────────
    const elements = this._parseLayoutPlaceholders(layoutXml, slideW, slideH, themeColors, masterXml);

    // ── 9. Thumbnail ─────────────────────────────────────
    let thumbnailData = null;
    for (const n of ['docProps/thumbnail.jpeg','docProps/thumbnail.jpg','docProps/thumbnail.png']) {
      const f = zip.file(n);
      if (f) { thumbnailData = await this._blobToDataUrl(await f.async('blob')); break; }
    }

    return {
      slideW, slideH,
      bgImageData:    fullSlideImage,  // slide completo (fundo + logos fixos)
      logoData:       null,             // sempre null: logo já está no bgImageData
      thumbnailData,
      elements,
      fullSlideImage,
    };
  },

  // ─────────────────────────────────────────────────────────────
  // renderFullSlide
  //
  // Renderiza o slide COMPLETO (fundo + todas as imagens fixas) num
  // único canvas e retorna como data URL PNG — SEM comprimir para JPEG.
  //
  // Este é o template_image que será salvo na variedade.
  // O CardRenderer depois apenas sobrepõe texto sobre esta imagem.
  //
  // Parâmetro outputW: largura em px do canvas (altura calculada pelo
  //                    aspect ratio do slide EMU).
  // ─────────────────────────────────────────────────────────────
  async renderFullSlide(zip, slideW, slideH, bgImageData, allImages, outputW = 1080) {
    // Calcula altura pelo aspect ratio EMU
    const aspect = slideH / slideW;
    const W = outputW;
    const H = Math.round(W * aspect);

    console.log(`[PptxParser.renderFullSlide] canvas=${W}×${H} | slideEMU=${slideW}×${slideH} | bgImageData=${bgImageData ? Math.round(bgImageData.length/1024)+'KB' : 'null'} | allImages=${allImages.length}`);

    const canvas = document.createElement('canvas');
    canvas.width  = W;
    canvas.height = H;
    const ctx = canvas.getContext('2d');

    // 1. Limpa com transparência (não preto!)
    ctx.clearRect(0, 0, W, H);

    // 2. Desenha fundo (background do master)
    if (bgImageData) {
      try {
        const bg = await this._loadImage(bgImageData);
        ctx.drawImage(bg, 0, 0, W, H);
        console.log(`[PptxParser.renderFullSlide] ✅ Fundo desenhado (${bg.naturalWidth}×${bg.naturalHeight})`);
      } catch (e) {
        console.warn('[PptxParser.renderFullSlide] ❌ Fundo não carregou:', e.message);
        // Sem fallback de cor sólida — deixa transparente para não esconder erro
      }
    } else {
      console.warn('[PptxParser.renderFullSlide] ⚠️ bgImageData é null — slide ficará sem fundo');
    }

    // 3. Desenha cada imagem dos outros slides/layouts em suas posições EMU
    //    allImages = [{ dataUrl, x, y, cx, cy, fname }]
    // Escala EMU → pixels do canvas
    // EMU = 914400 / 96 = 9525 EMU por pixel (96 DPI)
    const EMU = 914400 / 96;  // 9525
    // slideW em EMU → dividir por EMU → largura em pixels @ 96 DPI
    // sX: fator de escala de "pixels EMU" para "pixels canvas"
    const sX  = W / (slideW / EMU);
    const sY  = H / (slideH / EMU);

    let imgDrawn = 0;
    for (const img of allImages) {
      if (!img.dataUrl) continue;
      try {
        const image  = await this._loadImage(img.dataUrl);
        const px = Math.round((img.x  / EMU) * sX);
        const py = Math.round((img.y  / EMU) * sY);
        const pw = Math.round((img.cx / EMU) * sX);
        const ph = Math.round((img.cy / EMU) * sY);
        if (pw > 0 && ph > 0) {
          ctx.drawImage(image, px, py, pw, ph);
          imgDrawn++;
          console.log(`[PptxParser.renderFullSlide]   ✅ Imagem ${img.fname?.replace('ppt/media/','')} @ (${px},${py}) ${pw}×${ph}px`);
        } else {
          console.warn(`[PptxParser.renderFullSlide]   ⚠️ Imagem ${img.fname} com dimensão zero (pw=${pw} ph=${ph})`);
        }
      } catch (e) {
        console.warn(`[PptxParser.renderFullSlide]   ❌ Imagem ${img.fname} não carregou:`, e.message);
      }
    }

    console.log(`[PptxParser.renderFullSlide] ✅ Concluído: ${imgDrawn}/${allImages.length} imagens desenhadas`);

    // Retorna como PNG (preserva transparência)
    return canvas.toDataURL('image/png');
  },

  // ─── Descobre o arquivo de layout do slide ─────────────
  _findLayoutFile(slideRels) {
    const m = slideRels.match(/Type="[^"]*slideLayout"[^>]*Target="([^"]+)"/);
    if (m) {
      // "../slideLayouts/slideLayout5.xml" → "ppt/slideLayouts/slideLayout5.xml"
      return m[1].replace(/^\.\.\//, 'ppt/');
    }
    return 'ppt/slideLayouts/slideLayout1.xml';
  },

  // ─── Extrai TODAS as imagens (com posições EMU) de um XML + rels ─
  // Retorna array de { dataUrl, x, y, cx, cy, rId, fname }
  // Fontes: <p:pic> elements no XML com posição <xfrm>
  //
  // IMPORTANTE: DOMParser com 'text/xml' é namespace-aware.
  // querySelectorAll('pic') NÃO encontra <p:pic> com prefixo.
  // Usamos getElementsByTagNameNS com wildcard '*' para contornar isso.
  async _extractImagesWithPositions(zip, xml, relsText) {
    if (!xml || !relsText) return [];

    // Monta mapa rId → fname de media (a partir do .rels)
    const rIdToFile = {};
    const imgRels = [...relsText.matchAll(/Id="([^"]+)"[^>]*Target="[^"]*media\/([^"?#]+\.(png|jpg|jpeg|gif|bmp|webp))"/gi)];
    for (const rel of imgRels) {
      const rawName = rel[2].split('/').pop(); // pega só o nome do arquivo
      rIdToFile[rel[1]] = `ppt/media/${rawName}`;
    }

    if (Object.keys(rIdToFile).length === 0) return []; // sem imagens no rels

    const parser = new DOMParser();
    const doc    = parser.parseFromString(xml, 'text/xml');

    // Coleta todos os elementos <p:pic> independente do prefixo do namespace
    // getElementsByTagNameNS('*', 'pic') funciona com qualquer prefixo
    const pics = Array.from(doc.getElementsByTagNameNS('*', 'pic'));
    const images = [];

    for (const pic of pics) {
      // Posição EMU: busca <xfrm> dentro de <spPr>
      const spPrs = pic.getElementsByTagNameNS('*', 'spPr');
      const spPr  = spPrs[0];
      if (!spPr) continue;

      const xfrms = spPr.getElementsByTagNameNS('*', 'xfrm');
      const xfrm  = xfrms[0];
      if (!xfrm) continue;

      const offs = xfrm.getElementsByTagNameNS('*', 'off');
      const exts = xfrm.getElementsByTagNameNS('*', 'ext');
      const off  = offs[0];
      const ext  = exts[0];
      if (!off || !ext) continue;

      const x  = parseInt(off.getAttribute('x')  || '0');
      const y  = parseInt(off.getAttribute('y')  || '0');
      const cx = parseInt(ext.getAttribute('cx') || '0');
      const cy = parseInt(ext.getAttribute('cy') || '0');
      if (cx <= 0 || cy <= 0) continue;

      // rId da imagem: está no <a:blip r:embed="rIdX">
      const blips = pic.getElementsByTagNameNS('*', 'blip');
      const blip  = blips[0];
      if (!blip) continue;

      // Tenta todos os atributos que têm localName 'embed'
      let rId = '';
      const attrs = Array.from(blip.attributes);
      for (const attr of attrs) {
        if (attr.localName === 'embed') { rId = attr.value; break; }
      }
      if (!rId) continue;

      const fname = rIdToFile[rId];
      if (!fname) continue;

      const f = zip.file(fname);
      if (!f) continue;

      const blob    = await f.async('blob');
      const dataUrl = await this._blobToDataUrl(blob);
      images.push({ dataUrl, x, y, cx, cy, rId, fname });
    }

    return images;
  },

  // ─── Extrai imagem de fundo do slideMaster ─────────────
  // Estratégia 1: maior <p:pic> no master (imagem de fundo como objeto)
  // Estratégia 2: <p:bg> → <a:blipFill> → imagem de fundo via rels
  async _extractBgFromMaster(zip, masterRels) {
    // Monta mapa rId → fname
    const rIdToFile = {};
    const imgRels = [...masterRels.matchAll(/Id="([^"]+)"[^>]*Target="[^"]*media\/([^"?#]+\.(png|jpg|jpeg|gif|bmp|webp))"/gi)];
    for (const rel of imgRels) {
      const rawName = rel[2].split('/').pop();
      rIdToFile[rel[1]] = `ppt/media/${rawName}`;
    }

    if (Object.keys(rIdToFile).length === 0) return null;

    // Carrega todos os arquivos com seus tamanhos
    const fileEntries = [];
    for (const [rId, fname] of Object.entries(rIdToFile)) {
      const f = zip.file(fname);
      if (f) {
        const blob = await f.async('blob');
        fileEntries.push({ rId, fname, blob });
      }
    }
    if (fileEntries.length === 0) return null;

    // Tenta obter o rId do blipFill do fundo no XML do master
    const masterXml = await zip.file('ppt/slideMasters/slideMaster1.xml')?.async('text') || '';
    if (masterXml) {
      const parser = new DOMParser();
      const doc    = parser.parseFromString(masterXml, 'text/xml');

      // Busca <p:bg> → <p:bgPr> → <a:blipFill> → <a:blip r:embed="...">
      const bgs = doc.getElementsByTagNameNS('*', 'bg');
      for (const bg of Array.from(bgs)) {
        const blips = bg.getElementsByTagNameNS('*', 'blip');
        const blip  = blips[0];
        if (!blip) continue;
        const attrs = Array.from(blip.attributes);
        let rId = '';
        for (const attr of attrs) {
          if (attr.localName === 'embed') { rId = attr.value; break; }
        }
        if (rId && rIdToFile[rId]) {
          const f = zip.file(rIdToFile[rId]);
          if (f) {
            console.log('[PptxParser] Fundo via p:bg blipFill:', rIdToFile[rId]);
            return this._blobToDataUrl(await f.async('blob'));
          }
        }
      }
    }

    // Fallback: pega a maior imagem referenciada no master
    const best = fileEntries.sort((a, b) => b.blob.size - a.blob.size)[0];
    console.log('[PptxParser] Fundo via maior imagem do master:', best.fname);
    return this._blobToDataUrl(best.blob);
  },

  // ─── Extrai logo da variedade — busca em múltiplas fontes ────
  //
  //  ESTRATÉGIA (em ordem de preferência):
  //  1. slideLayout.rels  → imagem no próprio layout (logo real embutido)
  //  2. slide1.rels       → imagem no slide1 (logo pode estar só aqui)
  //  3. slideMaster.rels  → menor imagem no master (2ª imagem = logo)
  //
  //  Para a POSIÇÃO:
  //  a) Elemento <p:pic> com xfrm no layout
  //  b) SP placeholder type=pic com {{logo_variedade}} no layout
  //  c) Elemento <p:pic> com xfrm no slide1
  //  d) Null (CardRenderer usa posição proporcional como fallback)
  //
  async _extractLogo(zip, layoutXml, layoutRels, masterXml, masterRels, slide1Xml, slide1Rels, bgImageData) {
    const parser = new DOMParser();

    // ── Extrai posição do logo de um XML (layout ou slide) ─────
    const extractPosition = (xml) => {
      if (!xml) return null;
      const doc = parser.parseFromString(xml, 'text/xml');
      let pos = null;

      // 1a. PIC element (imagem embutida real) — namespace-safe
      const pics = Array.from(doc.getElementsByTagNameNS('*', 'pic'));
      for (const pic of pics) {
        if (pos) break;
        const spPrs = pic.getElementsByTagNameNS('*', 'spPr');
        const xfrms = spPrs[0]?.getElementsByTagNameNS('*', 'xfrm');
        const xfrm  = xfrms?.[0];
        const offs  = xfrm?.getElementsByTagNameNS('*', 'off');
        const exts  = xfrm?.getElementsByTagNameNS('*', 'ext');
        const off   = offs?.[0]; const ext = exts?.[0];
        if (off && ext) {
          pos = {
            x:  parseInt(off.getAttribute('x')  || '0'),
            y:  parseInt(off.getAttribute('y')  || '0'),
            cx: parseInt(ext.getAttribute('cx') || '0'),
            cy: parseInt(ext.getAttribute('cy') || '0'),
          };
        }
      }

      // 1b. SP placeholder type=pic ou com nome {{logo_variedade}}
      if (!pos) {
        const sps = Array.from(doc.getElementsByTagNameNS('*', 'sp'));
        for (const sp of sps) {
          if (pos) break;
          const cNvPrs = sp.getElementsByTagNameNS('*', 'cNvPr');
          const cNvPr  = cNvPrs[0];
          const nvPrs  = sp.getElementsByTagNameNS('*', 'nvPr');
          const phEls  = nvPrs[0]?.getElementsByTagNameNS('*', 'ph');
          const ph     = phEls?.[0];
          const name   = cNvPr?.getAttribute('name') || '';
          if ((ph?.getAttribute('type') === 'pic') || name.toLowerCase().includes('logo')) {
            const spPrs = sp.getElementsByTagNameNS('*', 'spPr');
            const xfrms = spPrs[0]?.getElementsByTagNameNS('*', 'xfrm');
            const xfrm  = xfrms?.[0];
            const offs  = xfrm?.getElementsByTagNameNS('*', 'off');
            const exts  = xfrm?.getElementsByTagNameNS('*', 'ext');
            const off   = offs?.[0]; const ext = exts?.[0];
            if (off && ext) {
              pos = {
                x:  parseInt(off.getAttribute('x')  || '0'),
                y:  parseInt(off.getAttribute('y')  || '0'),
                cx: parseInt(ext.getAttribute('cx') || '0'),
                cy: parseInt(ext.getAttribute('cy') || '0'),
              };
            }
          }
        }
      }

      return pos;
    };

    // ── Coleta imagens de um arquivo .rels ─────────────────────
    const collectImages = async (relsText, baseFolder) => {
      // Suporta tanto ../ quanto caminhos relativos simples
      const allImgRefs = [
        ...(relsText.matchAll(/Id="([^"]+)"[^>]+Target="\.\.\/media\/([^"]+\.(png|jpg|jpeg|gif|bmp|webp))"/gi)),
        ...(relsText.matchAll(/Id="([^"]+)"[^>]+Target="([^"]*media\/[^"]+\.(png|jpg|jpeg|gif|bmp|webp))"/gi)),
      ];
      const seen = new Set();
      const images = [];
      for (const ref of allImgRefs) {
        let fname = ref[2];
        if (!fname.startsWith('ppt/')) {
          fname = `ppt/media/${fname.replace(/^.*\/media\//, '')}`;
        }
        if (seen.has(fname)) continue;
        seen.add(fname);
        const f = zip.file(fname);
        if (f) {
          const blob = await f.async('blob');
          images.push({ fname, blob, rId: ref[1] });
        }
      }
      return images;
    };

    // ── FONTE 1: slideLayout ────────────────────────────────────
    const layoutImages = await collectImages(layoutRels, 'ppt/media/');
    if (layoutImages.length > 0) {
      // Pega a maior imagem do layout (logo)
      const biggest = layoutImages.sort((a, b) => b.blob.size - a.blob.size)[0];
      const dataUrl  = await this._blobToDataUrl(biggest.blob);
      const position = extractPosition(layoutXml);
      console.log(`[PptxParser] Logo encontrado no slideLayout: ${biggest.fname} (${Math.round(biggest.blob.size/1024)}KB)`);
      return { dataUrl, position };
    }

    // ── FONTE 2: slide1 ─────────────────────────────────────────
    const slide1Images = await collectImages(slide1Rels, 'ppt/media/');
    if (slide1Images.length > 0) {
      // Pega a maior imagem do slide1
      const biggest = slide1Images.sort((a, b) => b.blob.size - a.blob.size)[0];
      const dataUrl  = await this._blobToDataUrl(biggest.blob);
      // Tenta posição no slide1, senão usa layout
      const position = extractPosition(slide1Xml) || extractPosition(layoutXml);
      console.log(`[PptxParser] Logo encontrado no slide1: ${biggest.fname} (${Math.round(biggest.blob.size/1024)}KB)`);
      return { dataUrl, position };
    }

    // ── FONTE 3: slideMaster — imagens menores que o fundo ──────
    const masterImages = await collectImages(masterRels, 'ppt/media/');
    if (masterImages.length > 1) {
      // Ordena por tamanho: a MAIOR é o fundo, as menores podem ser logos
      const sorted = masterImages.sort((a, b) => b.blob.size - a.blob.size);
      // Remove a maior (fundo) e pega a próxima maior
      // Mas verifica se o bgImageData corresponde à maior do master para não confundir
      const bgDataUrl = bgImageData || '';
      let candidates = sorted.slice(1); // remove a maior (fundo)

      // Se bgImageData veio do master, a maior é o fundo — pega a 2ª
      if (candidates.length > 0) {
        const logo = candidates[0];
        const dataUrl = await this._blobToDataUrl(logo.blob);
        // Busca posição no master XML
        const position = extractPosition(masterXml) || extractPosition(layoutXml) || null;
        console.log(`[PptxParser] Logo encontrado no slideMaster (2ª maior imagem): ${logo.fname} (${Math.round(logo.blob.size/1024)}KB)`);
        return { dataUrl, position };
      }
    }

    // ── FONTE 4: slideMaster — única imagem pequena ─────────────
    // Se o master tem apenas 1 imagem mas ela é diferente do bg, tenta usá-la
    if (masterImages.length === 1) {
      const img = masterImages[0];
      // Só usa se for menor que 200KB (logos costumam ser pequenos)
      if (img.blob.size < 200 * 1024) {
        const dataUrl = await this._blobToDataUrl(img.blob);
        console.log(`[PptxParser] Logo encontrado no slideMaster (única imagem pequena): ${img.fname}`);
        return { dataUrl, position: null };
      }
    }

    console.warn('[PptxParser] Logo não encontrado em nenhuma fonte (layout, slide1, master).');
    return null;
  },

  // ─── Mantido para compatibilidade (delega para _extractLogo) ─
  async _extractLogoFromLayout(zip, layoutXml, layoutRels) {
    return this._extractLogo(zip, layoutXml, layoutRels, '', '', '', '', null);
  },

  // ─── Helper: busca elemento pelo localName, ignorando prefixo namespace ───
  // Equivalente a querySelector mas namespace-safe no DOMParser 'text/xml'
  _nsFind(el, localName) {
    if (!el) return null;
    const found = el.getElementsByTagNameNS('*', localName);
    return found.length ? found[0] : null;
  },
  _nsFindAll(el, localName) {
    if (!el) return [];
    return Array.from(el.getElementsByTagNameNS('*', localName));
  },

  // ─── Extrai mapa de cores do tema (schemeClr → hex) ────────
  async _extractThemeColors(zip) {
    const themeXml = await zip.file('ppt/theme/theme1.xml')?.async('text') || '';
    if (!themeXml) return {};
    const doc = new DOMParser().parseFromString(themeXml, 'text/xml');
    const scheme = doc.getElementsByTagNameNS('*', 'clrScheme')[0];
    if (!scheme) return {};
    const map = {};
    for (const child of Array.from(scheme.children)) {
      const name = child.localName; // dk1, lt1, dk2, lt2, accent1…
      const srgb = child.getElementsByTagNameNS('*', 'srgbClr')[0];
      const sys  = child.getElementsByTagNameNS('*', 'sysClr')[0];
      const hex  = srgb ? srgb.getAttribute('val')
                        : sys ? sys.getAttribute('lastClr') : null;
      if (hex) {
        map[name] = '#' + hex;
        // Aliases OOXML: dk1=tx1, lt1=bg1, dk2=tx2, lt2=bg2
        if (name === 'dk1') { map['tx1'] = '#' + hex; }
        if (name === 'lt1') { map['bg1'] = '#' + hex; map['lt1'] = '#' + hex; }
        if (name === 'dk2') { map['tx2'] = '#' + hex; }
        if (name === 'lt2') { map['bg2'] = '#' + hex; }
      }
    }
    console.log('[PptxParser] Tema cores:', JSON.stringify(map));
    return map;
  },

  // ─── Parse todos os shapes de texto do layout ──────────
  _parseLayoutPlaceholders(layoutXml, slideW, slideH, themeColors = {}, masterXml = '') {
    const parser = new DOMParser();
    const doc    = parser.parseFromString(layoutXml, 'text/xml');
    const elements = [];

    // Parse master para fallback de alinhamento
    let masterDoc = null;
    if (masterXml) {
      try { masterDoc = parser.parseFromString(masterXml, 'text/xml'); } catch {}
    }

    // Usa getElementsByTagNameNS para ser namespace-safe
    const sps = Array.from(doc.getElementsByTagNameNS('*', 'sp'));
    for (const sp of sps) {
      const el = this._parseLayoutShape(sp, slideW, slideH, themeColors, masterDoc);
      if (el) elements.push(el);
    }

    return elements;
  },

  _parseLayoutShape(sp, slideW, slideH, themeColors = {}, masterDoc = null) {
    const cNvPr = this._nsFind(sp, 'cNvPr');
    if (!cNvPr) return null;

    const name = cNvPr.getAttribute('name') || '';

    // Extrai todos os {{placeholder}} do nome
    const phMatches = [...name.matchAll(/\{\{(\w+)\}\}/g)];
    if (!phMatches.length) return null;

    const placeholders = phMatches.map(m => m[1].toLowerCase());

    // Ignora placeholders que são imagens (tratados separadamente)
    const imgPhs = ['logo_marca', 'logo_variedade', 'logo_cliente', 'logo', 'layout'];
    if (placeholders.every(p => imgPhs.includes(p))) return null;

    // Ignora se é ph type=pic
    const nvPr = this._nsFind(sp, 'nvPr');
    const phEl = nvPr ? this._nsFind(nvPr, 'ph') : null;
    if (phEl?.getAttribute('type') === 'pic') return null;

    // Posição: busca spPr → xfrm → off/ext
    const spPr = this._nsFind(sp, 'spPr');
    const xfrm = spPr ? this._nsFind(spPr, 'xfrm') : null;
    const off  = xfrm ? this._nsFind(xfrm, 'off') : null;
    const ext  = xfrm ? this._nsFind(xfrm, 'ext') : null;
    if (!off || !ext) return null;

    const x = parseInt(off.getAttribute('x')  || '0');
    const y = parseInt(off.getAttribute('y')  || '0');
    const w = parseInt(ext.getAttribute('cx') || '0');
    const h = parseInt(ext.getAttribute('cy') || '0');

    // ── Propriedades de estilo ────────────────────────────
    // Estratégia de resolução (ordem de prioridade, mais específico primeiro):
    //   1. txBody > a:p > a:pPr   → alinhamento real do parágrafo
    //   2. lstStyle > lvl5pPr     → estilo mais específico (nível 5)
    //   3. lstStyle > lvl4..1pPr  → estilos genéricos (fallback)
    // Para COR: percorre do lvl5 ao lvl1 e usa o PRIMEIRO nível que tiver solidFill.
    // Isso evita que lvl1pPr sem cor sobrescreva lvl5pPr com cor.

    const txBody   = this._nsFind(sp, 'txBody');
    const bodyPr   = txBody ? this._nsFind(txBody, 'bodyPr') : null;
    const lstStyle = this._nsFind(sp, 'lstStyle');

    // Helper: resolve cor de um elemento defRPr usando o tema real
    const resolveColor = (rpr) => {
      if (!rpr) return null;
      const solidFills = rpr.getElementsByTagNameNS('*', 'solidFill');
      if (!solidFills.length) return null;
      const sf   = solidFills[0];
      const srgb = sf.getElementsByTagNameNS('*', 'srgbClr')[0];
      if (srgb) return '#' + srgb.getAttribute('val');
      const scheme = sf.getElementsByTagNameNS('*', 'schemeClr')[0];
      if (scheme) {
        const key = scheme.getAttribute('val');
        // Usa o tema real extraído do PPTX; fallback para valores padrão Office
        return themeColors[key] || {
          'tx1':'#000000','dk1':'#000000','bg1':'#FFFFFF','lt1':'#FFFFFF',
          'tx2':'#0E2841','dk2':'#0E2841','bg2':'#E8E8E8','lt2':'#E8E8E8',
          'accent1':'#4472C4','accent2':'#ED7D31','accent3':'#A9D18E',
          'accent4':'#FFC000','accent5':'#5B9BD5','accent6':'#70AD47',
        }[key] || null;
      }
      return null;
    };

    let align     = null;
    let szRaw     = null;
    let boldAttr  = null;
    let fontFamilyRaw = null;
    let color     = null;

    // Prioridade 1: txBody > a:p > a:pPr (alinhamento do parágrafo real)
    // IMPORTANTE: Varre TODOS os parágrafos — prioriza o que tiver texto (placeholder),
    // pois parágrafos vazios podem ter algn incorreto.
    if (txBody) {
      const paras = Array.from(txBody.getElementsByTagNameNS('*', 'p'));
      // Primeiro tenta parágrafo que contenha texto/placeholder
      for (const para of paras) {
        const hasText = para.textContent.trim().length > 0;
        if (!hasText) continue;
        const pPr  = para.getElementsByTagNameNS('*', 'pPr')[0];
        const algn = pPr?.getAttribute('algn');
        if (algn && !align) {
          align = { l:'left', ctr:'center', r:'right', just:'left' }[algn] || 'left';
        }
      }
      // Se não encontrou em parágrafo com texto, usa qualquer parágrafo
      if (!align) {
        for (const para of paras) {
          const pPr  = para.getElementsByTagNameNS('*', 'pPr')[0];
          const algn = pPr?.getAttribute('algn');
          if (algn) {
            align = { l:'left', ctr:'center', r:'right', just:'left' }[algn] || 'left';
            break;
          }
        }
      }
    }

    // Prioridade 2: lstStyle — percorre do mais específico (lvl5) ao mais genérico (lvl1)
    // Para ALINHAMENTO: usa o primeiro nível que tiver algn
    // Para COR:         usa o primeiro nível que tiver solidFill (lvl5 primeiro!)
    // Para sz/bold/font: usa o primeiro nível que tiver o atributo
    if (lstStyle) {
      for (const lvl of ['lvl5pPr','lvl4pPr','lvl3pPr','lvl2pPr','lvl1pPr']) {
        const pPrs = lstStyle.getElementsByTagNameNS('*', lvl);
        const pPr  = pPrs.length ? pPrs[0] : null;
        if (!pPr) continue;

        // Alinhamento
        if (!align) {
          const algn = pPr.getAttribute('algn');
          if (algn) align = { l:'left', ctr:'center', r:'right', just:'left' }[algn] || 'left';
        }

        const rprEls = pPr.getElementsByTagNameNS('*', 'defRPr');
        const rpr    = rprEls.length ? rprEls[0] : null;
        if (!rpr) continue;

        // Cor — usa o PRIMEIRO nível que tiver solidFill (mais específico)
        if (!color) color = resolveColor(rpr);

        // Tamanho de fonte
        if (!szRaw) szRaw = rpr.getAttribute('sz');

        // Negrito
        if (boldAttr === null) {
          const b = rpr.getAttribute('b');
          if (b !== null) boldAttr = b;
        }

        // Família de fonte
        if (!fontFamilyRaw) {
          const lat = this._nsFind(rpr, 'latin');
          if (lat?.getAttribute('typeface')) fontFamilyRaw = lat.getAttribute('typeface');
        }
      }
    }

    // Fallback: txBody > rPr (runs do parágrafo)
    if (txBody && (!color || !szRaw || boldAttr === null || !fontFamilyRaw)) {
      const allRPr = Array.from(txBody.getElementsByTagNameNS('*', 'rPr'));
      for (const rPr of allRPr) {
        if (!color)          color = resolveColor(rPr);
        if (!szRaw)          szRaw = rPr.getAttribute('sz');
        if (boldAttr === null) { const b = rPr.getAttribute('b'); if (b !== null) boldAttr = b; }
        if (!fontFamilyRaw) {
          const lat = this._nsFind(rPr, 'latin');
          if (lat?.getAttribute('typeface')) fontFamilyRaw = lat.getAttribute('typeface');
        }
      }
    }

    // Prioridade 3: busca no slideMaster pelo mesmo nome de shape
    // Útil quando o layout não define algn/sz/color mas o master define
    if (masterDoc && (!align || !szRaw || !color)) {
      const masterSps = Array.from(masterDoc.getElementsByTagNameNS('*', 'sp'));
      for (const msp of masterSps) {
        const mCnvPr = this._nsFind(msp, 'cNvPr');
        if (!mCnvPr) continue;
        const mName = mCnvPr.getAttribute('name') || '';
        // Verifica se este shape do master tem o mesmo placeholder
        const mPhs = [...mName.matchAll(/\{\{(\w+)\}\}/g)].map(m => m[1].toLowerCase());
        if (!mPhs.some(p => placeholders.includes(p))) continue;

        const mTxBody   = this._nsFind(msp, 'txBody');
        const mLstStyle = this._nsFind(msp, 'lstStyle');

        // Alinhamento do master
        if (!align && mTxBody) {
          const mParas = Array.from(mTxBody.getElementsByTagNameNS('*', 'p'));
          for (const para of mParas) {
            if (!para.textContent.trim()) continue;
            const pPr = para.getElementsByTagNameNS('*', 'pPr')[0];
            const algn = pPr?.getAttribute('algn');
            if (algn) { align = { l:'left', ctr:'center', r:'right', just:'left' }[algn] || 'left'; break; }
          }
        }
        if (!align && mLstStyle) {
          for (const lvl of ['lvl5pPr','lvl4pPr','lvl3pPr','lvl2pPr','lvl1pPr']) {
            const pPr = mLstStyle.getElementsByTagNameNS('*', lvl)[0];
            const algn = pPr?.getAttribute('algn');
            if (algn && !align) { align = { l:'left', ctr:'center', r:'right', just:'left' }[algn] || 'left'; }
          }
        }
        break;
      }
    }

    // Valores finais
    if (!align)       align = 'left';
    if (!color)       color = '#000000';
    const fontSizePt  = szRaw ? parseInt(szRaw) / 100 : 12;
    const bold        = boldAttr === '1';
    const fontFamily  = fontFamilyRaw || 'Arial';

    // Padding interno
    const lIns = parseInt(bodyPr?.getAttribute('lIns') || '91440');
    const rIns = parseInt(bodyPr?.getAttribute('rIns') || '91440');
    const tIns = parseInt(bodyPr?.getAttribute('tIns') || '45720');
    const bIns = parseInt(bodyPr?.getAttribute('bIns') || '45720');

    // Campo composto ex: {{cidade}}/{{estado}}
    // Campos de produtividade e unidade são sempre negrito (regra visual do card)
    const BOLD_FIELDS = ['produtividade_int', 'produtividade_dec', 'unidade'];
    const isBoldField = placeholders.some(p => BOLD_FIELDS.includes(p));
    const finalBold   = isBoldField ? true : bold;

    // Regra de alinhamento inteligente para produtividade_int:
    // O número inteiro (ex: "90") fica alinhado à DIREITA dentro do seu box no PPTX.
    // O layout/master frequentemente não define algn explicitamente para este campo,
    // então o parser retorna 'left' por padrão — o que fica incorreto.
    // Heurística: se o align ainda for 'left' e o placeholder for produtividade_int,
    // e o box for largo (> 30% do slide) mas começar antes de 20%, é alinhado à direita.
    if (align === 'left' && placeholders.includes('produtividade_int')) {
      const wPct = (w / slideW) * 100;
      const xPct = (x / slideW) * 100;
      if (wPct > 30 && xPct < 20) {
        align = 'right';
      }
    }

    const result = placeholders.length > 1 ? {
      placeholder: 'cidade_estado',
      placeholders,
      separator: '/',
      x, y, w, h,
      fontSizePt, bold: finalBold, italic: false, fontFamily, color, align,
      lIns, rIns, tIns, bIns, name,
    } : {
      placeholder: placeholders[0],
      x, y, w, h,
      fontSizePt, bold: finalBold, italic: false, fontFamily, color, align,
      lIns, rIns, tIns, bIns, name,
    };

    // Log de diagnóstico para campos de produtividade
    const PROD_FIELDS = ['produtividade_int','produtividade_dec','unidade'];
    if (placeholders.some(p => PROD_FIELDS.includes(p))) {
      const slideWref = 6858000;
      console.log(`[parseShape] ${placeholders.join('+')}`, {
        x, xPct: (x/slideWref*100).toFixed(1)+'%',
        w, wPct: (w/slideWref*100).toFixed(1)+'%',
        align, lIns, rIns,
        fontSizePt,
      });
    }

    return result;
  },

  _blobToDataUrl(blob) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload  = e => res(e.target.result);
      r.onerror = rej;
      r.readAsDataURL(blob);
    });
  },

  _loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
      const s = document.createElement('script');
      s.src = src; s.onload = resolve;
      s.onerror = () => reject(new Error('Falha ao carregar: ' + src));
      document.head.appendChild(s);
    });
  },

  // ─── Carrega uma imagem a partir de uma URL ou data URL ─────
  // Necessário para renderFullSlide (PptxParser não usa CardRenderer)
  _loadImage(src) {
    return new Promise((resolve, reject) => {
      if (!src) return reject(new Error('Sem imagem (src vazio)'));
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload  = () => resolve(img);
      img.onerror = () => reject(new Error('Falha ao carregar imagem: ' + String(src).substring(0, 80)));
      img.src = src;
    });
  },
};


/* ─────────────────────────────────────────────────────────────
   CARD RENDERER v8
   Renderiza o card no canvas com alta resolução.
   Preserva posições e fontes do PPTX original.
───────────────────────────────────────────────────────────── */
const CardRenderer = {

  /**
   * Renderiza um card completo num canvas.
   *
   * ESTRATÉGIA v9 (layout preservado):
   * - bgImageData = PNG do slide completo (fundo + logos já compostos)
   *   O CardRenderer NÃO reinjeta logos — eles já estão no bgImageData.
   * - logoData = ignorado (mantido apenas para compatibilidade com legado)
   * - elements = apenas placeholders de TEXTO a sobrepor
   *
   * @param {object} params
   *   bgImageData  – PNG do slide completo (fundo + logos fixos)
   *   logoData     – ignorado na v9 (logo já está no bgImageData)
   *   elements     – array de elementos de texto do PPTX
   *   data         – mapa de dados { produtividade_int, data_plantio, ... }
   *   slideW/slideH – dimensões EMU do PPTX original
   *   outputW      – largura de saída em px
   * @returns {HTMLCanvasElement}
   */
  async render({ bgImageData, logoData, elements, data, slideW, slideH, outputW }) {
    // 1. Carrega o template_image (slide completo com logos preservados)
    const bgImg = await this._loadImage(bgImageData);

    const W = outputW || bgImg.naturalWidth;
    const H = Math.round(W * (bgImg.naturalHeight / bgImg.naturalWidth));

    const canvas = document.createElement('canvas');
    canvas.width  = W;
    canvas.height = H;
    const ctx = canvas.getContext('2d');

    // 2. Desenha o slide completo (fundo + logos fixos já compostos)
    ctx.drawImage(bgImg, 0, 0, W, H);

    // 3. Sobrepõe APENAS os placeholders de texto
    //    Ignora quaisquer elementos do tipo logo_image (não mais usados na v9)
    for (const el of (elements || [])) {
      if (el.type === 'logo_image') continue; // legado — ignorar
      await this._drawElement(ctx, el, data, slideW, slideH, W, H);
    }

    return canvas;
  },

  async _drawElement(ctx, el, data, slideW, slideH, canvasW, canvasH) {
    // Escala EMU → px do canvas
    const EMU = 914400 / 96; // EMU por pixel (96 DPI)
    const scaleX = canvasW / (slideW / EMU);
    const scaleY = canvasH / (slideH / EMU);
    const emuX = emu => emu / EMU * scaleX;
    const emuY = emu => emu / EMU * scaleY;

    // LOG diagnóstico para campos de produtividade (remover após debug)
    const isProdField = ['produtividade_int','produtividade_dec','unidade'].includes(el.placeholder);
    if (isProdField) {
      console.log(`[_drawElement] ${el.placeholder}`, {
        x: el.x, y: el.y, w: el.w, h: el.h,
        xPct: (el.x/slideW*100).toFixed(1)+'%',
        yPct: (el.y/slideH*100).toFixed(1)+'%',
        wPct: (el.w/slideW*100).toFixed(1)+'%',
        align: el.align,
        lIns: el.lIns, rIns: el.rIns,
        fontSizePt: el.fontSizePt,
        bold: el.bold,
        color: el.color,
        canvasW, canvasH,
        scaleX: scaleX.toFixed(3), scaleY: scaleY.toFixed(3),
        bx: emuX(el.x).toFixed(1),
        bw: emuX(el.w).toFixed(1),
        lPad: emuX(el.lIns||91440).toFixed(1),
      });
    }

    // Resolve o valor do placeholder
    let text = '';
    if (el.placeholders) {
      // Campo composto (ex: cidade/estado)
      text = el.placeholders
        .map(ph => data[ph] || '')
        .filter(Boolean)
        .join(el.separator || '/');
    } else {
      text = (data[el.placeholder] !== undefined && data[el.placeholder] !== null)
        ? String(data[el.placeholder])
        : '';
    }

    if (!text) return;

    const bx = emuX(el.x);
    const by = emuY(el.y);
    const bw = emuX(el.w);
    const bh = emuY(el.h);

    ctx.save();

    // Escala o tamanho da fonte proporcionalmente ao canvas
    // O tamanho nominal é em pt (1pt = 1.333px a 96dpi)
    // Mas usamos scaleY para manter proporção com o slide
    const nominalPx = el.fontSizePt * (96 / 72); // pt → px a 96dpi
    const scaledPx  = nominalPx * scaleY;

    // Padding interno
    const lPad = emuX(el.lIns || 91440);
    const rPad = emuX(el.rIns || 91440);
    const tPad = emuY(el.tIns || 45720);
    const usableW = Math.max(1, bw - lPad - rPad);

    // Fonte
    const weight = el.bold   ? 'bold' : 'normal';
    const style  = el.italic ? 'italic ' : '';
    const family = this._safeFont(el.fontFamily);

    // Auto-shrink: reduz fonte se o texto não couber na largura
    let fsPx = scaledPx;
    ctx.font = `${style}${weight} ${fsPx}px ${family}`;
    const naturalW = ctx.measureText(text).width;
    if (naturalW > usableW && usableW > 0) {
      fsPx = Math.max(6, fsPx * (usableW / naturalW) * 0.97);
    }

    ctx.font         = `${style}${weight} ${fsPx.toFixed(2)}px ${family}`;
    ctx.fillStyle    = el.color || '#FFFFFF';
    ctx.textBaseline = 'alphabetic';   // baseline padrão — mais previsível entre fontes
    ctx.shadowColor  = 'transparent';
    ctx.shadowBlur   = 0;

    // ── Override de alinhamento em render time ────────────────────────
    // O banco pode ter align='left' salvo de uma versão antiga do parser.
    // Aplicamos as mesmas heurísticas do _parseLayoutShape para corrigir.
    let effectiveAlign = el.align || 'left';

    // produtividade_int: se box > 30% de largura e começa antes de 20%,
    // o número deve ser alinhado à DIREITA (faz o "colinho" antes do decimal)
    if (effectiveAlign === 'left' && el.placeholder === 'produtividade_int') {
      const wPctRt = (el.w / slideW) * 100;
      const xPctRt = (el.x / slideW) * 100;
      if (wPctRt > 30 && xPctRt < 20) {
        effectiveAlign = 'right';
        if (isProdField) console.log('[_drawElement] override align → right para produtividade_int');
      }
    }

    // Posição X conforme alinhamento
    let drawX;
    if (effectiveAlign === 'center') {
      ctx.textAlign = 'center';
      drawX = bx + bw / 2;
    } else if (effectiveAlign === 'right') {
      ctx.textAlign = 'right';
      drawX = bx + bw - rPad;
    } else {
      ctx.textAlign = 'left';
      drawX = bx + lPad;
    }

    // ── Posição Y usando métricas reais da fonte ──────────────────────
    // measureText().actualBoundingBoxAscent dá a distância do baseline ao topo do glifo.
    // Usar 'alphabetic' baseline + ascent real é muito mais preciso do que fsPx * 1.2
    // (que varia bastante entre fontes e causa deslocamento em mobile vs desktop).
    const metrics = ctx.measureText(text);

    // Ascent real: distância do baseline até o topo do glifo
    const ascent  = (metrics.actualBoundingBoxAscent  > 0)
                    ? metrics.actualBoundingBoxAscent
                    : fsPx * 0.78;
    const descent = (metrics.actualBoundingBoxDescent > 0)
                    ? metrics.actualBoundingBoxDescent
                    : fsPx * 0.22;
    const realH   = ascent + descent;  // altura real do bloco de texto

    // Centraliza o glifo verticalmente dentro do box:
    // Se o box é maior que o glifo → centralizar; se menor → alinhar ao topo + tPad
    let drawY;
    if (bh >= realH) {
      // Box maior: centralizar verticalmente, com tPad mínimo
      const topOfGlyph = by + Math.max(tPad, (bh - realH) / 2);
      drawY = topOfGlyph + ascent;  // baseline = topo do glifo + ascent
    } else {
      // Box menor que o glifo (fonte muito grande): alinhar ao topo + tPad
      drawY = by + tPad + ascent;
    }

    ctx.fillText(text, drawX, drawY);
    ctx.restore();
  },

  /** Mapeia HelveticaNeueLT Pro → Barlow (Google Fonts) — idêntico ao CardGenerator._safeFont */
  _safeFont(family) {
    if (!family) return '"Barlow", Arial, sans-serif';
    if (/HelveticaNeueLT.*(?:Cn|BlkEx|HvCn|BdCn|MdCn|LtCn)/i.test(family))
      return '"Barlow Condensed", "Helvetica Neue", Helvetica, Arial, sans-serif';
    if (/HelveticaNeueLT.*(?:Ex|SemiCn)/i.test(family))
      return '"Barlow Semi Condensed", "Helvetica Neue", Helvetica, Arial, sans-serif';
    if (family.includes('HelveticaNeueLT') || family.includes('Helvetica'))
      return '"Barlow", "Helvetica Neue", Helvetica, Arial, sans-serif';
    return `"${family}", Arial, sans-serif`;
  },

  _loadImage(src) {
    return new Promise((resolve, reject) => {
      if (!src) return reject(new Error('Sem imagem'));
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload  = () => resolve(img);
      img.onerror = () => reject(new Error('Falha ao carregar imagem'));
      img.src = src;
    });
  },
};


/* ─────────────────────────────────────────────────────────────
   CARD GENERATOR v8
   Modal de preview, geração e download de cards
───────────────────────────────────────────────────────────── */
const CardGenerator = {
  currentRecord: null,
  _varietyCache: {},

  /**
   * Mapeamento HelveticaNeueLT Pro → Barlow (Google Fonts)
   * Usado em _safeFont e _ensureFontsLoaded.
   */
  _HELVETICA_MAP: {
    'HelveticaNeueLT Pro 47 LtCn':  '"Barlow Condensed", sans-serif',
    'HelveticaNeueLT Pro 57 Cn':    '"Barlow Condensed", sans-serif',
    'HelveticaNeueLT Pro 67 MdCn':  '"Barlow Condensed", sans-serif',
    'HelveticaNeueLT Pro 77 BdCn':  '"Barlow Condensed", sans-serif',
    'HelveticaNeueLT Pro 87 HvCn':  '"Barlow Condensed", sans-serif',
    'HelveticaNeueLT Pro 93 BlkEx': '"Barlow Condensed", sans-serif',
    'HelveticaNeueLT Pro 53 Ex':    '"Barlow Semi Condensed", sans-serif',
    'HelveticaNeueLT Pro 63 MdEx':  '"Barlow Semi Condensed", sans-serif',
    'HelveticaNeueLT Pro 73 BdEx':  '"Barlow Semi Condensed", sans-serif',
    'HelveticaNeueLT Pro 45 Lt':    '"Barlow", sans-serif',
    'HelveticaNeueLT Pro 55 Roman': '"Barlow", sans-serif',
    'HelveticaNeueLT Pro 65 Md':    '"Barlow", sans-serif',
    'HelveticaNeueLT Pro 75 Bd':    '"Barlow", sans-serif',
    'HelveticaNeueLT Pro 85 Hv':    '"Barlow", sans-serif',
    'HelveticaNeueLT Pro 95 Blk':   '"Barlow", sans-serif',
  },

  /** Mapeia HelveticaNeueLT Pro → Barlow (Google Fonts) */
  _safeFont(family) {
    if (!family) return '"Barlow", Arial, sans-serif';
    if (this._HELVETICA_MAP[family]) return this._HELVETICA_MAP[family];
    if (/HelveticaNeueLT.*(?:Cn|BlkEx|HvCn|BdCn|MdCn|LtCn)/i.test(family))
      return '"Barlow Condensed", "Helvetica Neue", Helvetica, Arial, sans-serif';
    if (/HelveticaNeueLT.*(?:Ex|SemiCn)/i.test(family))
      return '"Barlow Semi Condensed", "Helvetica Neue", Helvetica, Arial, sans-serif';
    if (family.includes('HelveticaNeueLT') || family.includes('Helvetica'))
      return '"Barlow", "Helvetica Neue", Helvetica, Arial, sans-serif';
    return `"${family}", Arial, sans-serif`;
  },

  /** Garante que as fontes Barlow estejam carregadas antes de renderizar */
  async _ensureFontsLoaded() {
    if (!document.fonts) return;
    try {
      await Promise.all([
        '400', '700', '800', '900', '300'
      ].flatMap(w => [
        document.fonts.load(`${w} 16px "Barlow Condensed"`).catch(() => {}),
        document.fonts.load(`${w} 16px "Barlow"`).catch(() => {}),
        document.fonts.load(`${w} 16px "Barlow Semi Condensed"`).catch(() => {}),
      ]));
    } catch {}
  },

  /* ═══════════════════════════════════════════════════
     OPEN PREVIEW
  ═══════════════════════════════════════════════════ */
  async openPreview(record) {
    this.currentRecord = record;

    const modal = document.getElementById('cardModal');
    if (!modal) return;
    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';

    const canvasDiv   = document.getElementById('cardCanvas');
    const downloadBtn = document.getElementById('btnDownloadCard');
    const editBtn     = document.getElementById('btnEditRecord');

    // Oculta botão de edição para usuários comuns
    if (editBtn) {
      const isAdmin = typeof AccessControl !== 'undefined' ? AccessControl.isAdmin() : true;
      editBtn.style.display = isAdmin ? '' : 'none';
    }

    canvasDiv.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;
      height:200px;color:#888;font-size:13px;gap:8px">
      <span class="loading"></span> Preparando card...</div>`;
    if (downloadBtn) downloadBtn.disabled = true;

    const variety = await this._getVariety(record.variety_id);

    if (!variety?.template_image) {
      canvasDiv.innerHTML = `
        <div class="card-no-template">
          <i class="fas fa-image" style="font-size:48px;color:#ccc;margin-bottom:12px"></i>
          <h3 style="margin:0 0 8px;font-size:16px;color:#555">Nenhum modelo vinculado</h3>
          <p style="font-size:13px;color:#888;text-align:center;max-width:240px">
            Acesse <strong>Variedades</strong> e faça o upload do PPTX para
            <strong>${this._esc(record.variety_name || 'esta variedade')}</strong>.
          </p>
        </div>`;
    } else {
      if (downloadBtn) downloadBtn.disabled = false;
      await this._renderPreview(canvasDiv, record, variety, 360);
    }

    if (downloadBtn) downloadBtn.onclick = () => this.download();

    document.getElementById('btnEditRecord')?.addEventListener('click', () => {
      this.closeModal();
      if (typeof User !== 'undefined' && User.editRecord) User.editRecord(record.id);
    }, { once: true });

    const closeBtn = document.getElementById('btnCloseModal');
    if (closeBtn) closeBtn.onclick = () => this.closeModal();

    const overlay = document.getElementById('modalOverlay');
    if (overlay) overlay.onclick = () => this.closeModal();
  },

  closeModal() {
    const modal = document.getElementById('cardModal');
    if (modal) modal.classList.add('hidden');
    document.body.style.overflow = '';
  },

  /* ═══════════════════════════════════════════════════
     RENDER PREVIEW (baixa resolução para exibição)
  ═══════════════════════════════════════════════════ */
  async _renderPreview(wrapper, record, variety, displayW) {
    try {
      const canvas = await this._buildCanvas(record, variety, displayW);
      if (!canvas) return;
      canvas.style.cssText = `width:${Math.min(displayW, 360)}px;max-width:100%;display:block;
        border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,.3);margin:0 auto;`;
      if (wrapper) { wrapper.innerHTML = ''; wrapper.appendChild(canvas); }
    } catch (err) {
      console.error('renderPreview:', err);
      if (wrapper) wrapper.innerHTML = `<div class="card-no-template" style="color:#c00">
        <i class="fas fa-exclamation-circle"></i>
        <p>Erro ao renderizar card.<br><small>${err.message}</small></p></div>`;
    }
  },

  /* ═══════════════════════════════════════════════════
     BUILD CANVAS — núcleo da renderização (v9)
     - template_image = slide completo (fundo + logos fixos)
     - Apenas texto é sobreposto pelo CardRenderer
  ═══════════════════════════════════════════════════ */
  async _buildCanvas(record, variety, outputW) {
    // ── Garante que as fontes estejam carregadas ───────
    await this._ensureFontsLoaded();

    // ── Mapa de dados ──────────────────────────────────
    const data = this._buildDataMap(record, variety);

    // ── Modo PPTX com elementos pré-processados ────────
    if (variety.pptx_elements && variety.pptx_elements !== '[]' && variety.pptx_elements !== '') {
      try {
        const elements = typeof variety.pptx_elements === 'string'
          ? JSON.parse(variety.pptx_elements)
          : variety.pptx_elements;

        console.log('[CardGenerator] Modo PPTX v9:', {
          hasBg:    !!variety.template_image,
          elements: elements.length,
          slideW:   variety.pptx_slide_w || 6858000,
          slideH:   variety.pptx_slide_h || 12193588,
        });

        // Log dos elementos de produtividade salvos no banco
        const TARGET = ['produtividade_int','produtividade_dec','unidade'];
        const sW = variety.pptx_slide_w || 6858000;
        const sH = variety.pptx_slide_h || 12193588;
        const prodEls = elements.filter(e => TARGET.includes(e.placeholder));
        prodEls.forEach(e => {
          console.log(`[pptx_elements DB] ${e.placeholder}`, {
            x: e.x, xPct: (e.x/sW*100).toFixed(1)+'%',
            y: e.y, yPct: (e.y/sH*100).toFixed(1)+'%',
            w: e.w, wPct: (e.w/sW*100).toFixed(1)+'%',
            h: e.h, hPct: (e.h/sH*100).toFixed(1)+'%',
            align: e.align, lIns: e.lIns, rIns: e.rIns,
            fontSizePt: e.fontSizePt, bold: e.bold, color: e.color,
          });
        });
        // Linha compacta para fácil cópia
        if (prodEls.length) {
          console.log('📋 COPIE ESTA LINHA →', JSON.stringify(prodEls.map(e=>({
            ph:e.placeholder, x:(e.x/sW*100).toFixed(1), y:(e.y/sH*100).toFixed(1),
            w:(e.w/sW*100).toFixed(1), h:(e.h/sH*100).toFixed(1),
            align:e.align, lIns:e.lIns, rIns:e.rIns, fs:e.fontSizePt
          }))));
        }

        // CardRenderer.render: desenha template_image (slide completo) + sobrepõe texto
        // logoData = null → logo já está no template_image, não reinjetar
        return await CardRenderer.render({
          bgImageData: variety.template_image,
          logoData:    null,              // v9: logo fixo no template_image
          elements,
          data,
          slideW: variety.pptx_slide_w || 6858000,
          slideH: variety.pptx_slide_h || 12193588,
          outputW,
        });
      } catch (e) {
        console.warn('[CardGenerator] Falha no modo PPTX, usando legado:', e.message);
      }
    }

    // ── Modo legado: coordenadas calibradas manualmente ─
    return this._buildCanvasLegacy(record, variety, outputW, data);
  },

  /* ═══════════════════════════════════════════════════
     CONSTRUÇÃO DO MAPA DE DADOS
     Regras de negócio: separação da produtividade,
     formatação cidade/estado, etc.
  ═══════════════════════════════════════════════════ */
  _buildDataMap(record, variety) {
    // ── Separação automática da produtividade ──────────
    // Aceita tanto vírgula (187,1) quanto ponto (187.1) como separador decimal
    let valor = (record.productivity || '').toString().trim();
    // Normaliza: se veio com ponto decimal (do input number), converte ponto → vírgula
    if (valor.includes('.') && !valor.includes(',')) {
      valor = valor.replace('.', ',');
    }
    const partes = valor.split(',');
    const produtividade_int = partes[0] || '';
    const produtividade_dec = partes[1] !== undefined ? ',' + partes[1] : '';

    return {
      // Produtividade
      produtividade:     valor,
      produtividade_int,
      produtividade_dec,

      // Unidade
      unidade:           record.unit || '',

      // Datas
      data_plantio:      record.planting_date  || '',
      data_colheita:     record.harvest_date   || '',

      // Localização
      cidade_estado:     [record.city, record.state].filter(Boolean).join('/') || '',
      cidade:            record.city  || '',
      estado:            record.state || '',

      // Produtor / Fazenda
      produtor:          record.producer_name || '',
      fazenda:           record.farm_name     || '',

      // Área
      area:              record.area || '',

      // Variedade
      variedade:         record.variety_name || variety?.name || '',
      cultura:           record.culture      || variety?.culture || '',
      tecnologia:        record.technology   || variety?.technology || '',
      safra:             record.season       || '',
    };
  },

  /* ═══════════════════════════════════════════════════
     DOWNLOAD — alta resolução PNG + upload OneDrive
  ═══════════════════════════════════════════════════ */
  async download() {
    const btn      = document.getElementById('btnDownloadCard');
    const origHtml = btn?.innerHTML || '';
    if (btn) { btn.innerHTML = '<span class="loading"></span> Gerando…'; btn.disabled = true; }

    try {
      const record  = this.currentRecord;
      if (!record) throw new Error('Nenhum registro selecionado');

      const variety = await this._getVariety(record.variety_id);
      if (!variety?.template_image) {
        App.Toast.show('Nenhum template vinculado à variedade.', 'error');
        return;
      }

      // Alta resolução: 3× o tamanho de preview (720px × 3 = 2160px)
      const bgImg  = await CardRenderer._loadImage(variety.template_image);
      const hiresW = Math.min(bgImg.naturalWidth * 3, 2160);

      const canvas = await this._buildCanvas(record, variety, hiresW);
      if (!canvas) return;

      // Gera PNG de alta qualidade (não JPEG)
      const pngDataUrl = canvas.toDataURL('image/png');

      // Nome padronizado do arquivo
      const cardFilename = typeof OneDrive !== 'undefined'
        ? OneDrive.buildFilename(record, 'png')
        : this._buildFilename(record, '.png');

      // ── Download local ──────────────────────────────
      const link = document.createElement('a');
      link.href     = pngDataUrl;
      link.download = cardFilename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      // ── Upload para OneDrive (assíncrono) ───────────
      let cardOnedrivePath = '';
      let cardOnedriveId   = '';
      let termoOnedrivePath = '';
      let termoOnedriveId  = '';

      if (typeof OneDrive !== 'undefined' && await OneDrive.isEnabled()) {
        btn.innerHTML = '<span class="loading"></span> Enviando ao OneDrive…';

        try {
          // Upload do card
          const cardRes = await OneDrive.uploadCard(pngDataUrl, record, cardFilename);
          if (cardRes) {
            cardOnedrivePath = cardRes.path || '';
            cardOnedriveId   = cardRes.id   || '';
            App.Toast.show('✅ Card enviado ao OneDrive!', 'success');
          }

          // Upload do termo (se existir no registro)
          if (record.termo_file && record.termo_nome_padronizado) {
            const termoRes = await OneDrive.uploadTermo(
              this._dataUrlToFile(record.termo_file, record.termo_filename || 'termo.pdf'),
              record,
              record.termo_nome_padronizado
            );
            if (termoRes) {
              termoOnedrivePath = termoRes.path || '';
              termoOnedriveId   = termoRes.id   || '';
            }
          }
        } catch (odErr) {
          console.warn('[CardGenerator] OneDrive upload falhou:', odErr);
          App.Toast.show('⚠️ OneDrive indisponível. Card salvo localmente.', 'warning');
        }
      }

      // ── Atualiza registro com dados do OneDrive ─────
      const updateData = {
        status:             'published',
        card_filename:      cardFilename,
        card_onedrive_path: cardOnedrivePath,
        card_onedrive_id:   cardOnedriveId,
      };
      if (termoOnedrivePath) {
        updateData.termo_onedrive_path = termoOnedrivePath;
        updateData.termo_onedrive_id   = termoOnedriveId;
      }

      if (record?.id) {
        try {
          await API.updateRecord(record.id, updateData);
          // Atualiza o record local
          Object.assign(record, updateData);
        } catch {}
      }

      // ── Log de auditoria ────────────────────────────
      if (typeof OneDrive !== 'undefined') {
        OneDrive.log('card_generated', record, {
          card_filename: cardFilename,
          card_path:     cardOnedrivePath,
          card_id:       cardOnedriveId,
          termo_filename: record.termo_nome_padronizado || '',
          termo_path:    termoOnedrivePath,
          termo_id:      termoOnedriveId,
        });
      }

      App.Toast.show('✅ Card baixado com sucesso!', 'success');

      // ── Fecha o modal e exibe tela de sucesso ───────
      this.closeModal();

      // Passa para a tela de sucesso se usuário comum
      if (typeof User !== 'undefined' && typeof AccessControl !== 'undefined' && !AccessControl.isAdmin()) {
        const previewDataUrl = canvas.toDataURL('image/jpeg', 0.75); // thumbnail para preview
        setTimeout(() => User.showSuccessScreen(record, previewDataUrl), 200);
      }

    } catch (err) {
      console.error('CardGenerator.download:', err);
      App.Toast.show('Erro ao gerar card. Tente novamente.', 'error');
    } finally {
      if (btn) { btn.innerHTML = origHtml; btn.disabled = false; }
    }
  },

  /** Converte dataUrl de volta para File */
  _dataUrlToFile(dataUrl, filename) {
    if (!dataUrl) return null;
    try {
      const [header, data] = dataUrl.split(',');
      const mime   = header.match(/:(.*?);/)[1];
      const binary = atob(data);
      const array  = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) array[i] = binary.charCodeAt(i);
      return new File([array], filename, { type: mime });
    } catch { return null; }
  },

  /* ═══════════════════════════════════════════════════
     MODO LEGADO — coordenadas calibradas manualmente
  ═══════════════════════════════════════════════════ */
  async _buildCanvasLegacy(record, variety, outputW, data) {
    let bgImg;
    try {
      bgImg = await CardRenderer._loadImage(variety.template_image);
    } catch {
      return null;
    }

    const W = outputW || bgImg.naturalWidth;
    const H = Math.round(W * (bgImg.naturalHeight / bgImg.naturalWidth));

    const canvas = document.createElement('canvas');
    canvas.width  = W;
    canvas.height = H;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(bgImg, 0, 0, W, H);

    const coords     = this._parseCoords(variety.field_coords);
    const varColor   = variety.primary_color || '#3CB226';

    this._drawFieldsLegacy(ctx, record, W, H, coords, varColor);
    return canvas;
  },

  DEFAULT_COORDS: {
    planting_date: { x: 0.26, y: 0.420, align: 'center', size: 0.032, weight: '700', color: '#2E4A1E' },
    harvest_date:  { x: 0.72, y: 0.420, align: 'center', size: 0.032, weight: '700', color: '#2E4A1E' },
    productivity:  { x: 0.50, y: 0.555, align: 'center', size: 0.175, weight: '900', color: 'variety' },
    unit:          { x: 0.50, y: 0.610, align: 'center', size: 0.038, weight: '700', color: '#4A6741' },
    location:      { x: 0.50, y: 0.678, align: 'center', size: 0.036, weight: '700', color: '#FFFFFF' },
    producer_name: { x: 0.50, y: 0.728, align: 'center', size: 0.034, weight: '700', color: '#1A2E12' },
    farm_name:     { x: 0.50, y: 0.760, align: 'center', size: 0.026, weight: '400', color: '#4A6741' },
    area_badge:    { x: 0.50, y: 0.812, align: 'center', size: 0.028, weight: '700', color: '#FFFFFF', badge: true }
  },

  _drawFieldsLegacy(ctx, record, W, H, coords, varietyColor) {
    const C   = coords;
    const px  = p => p * W;
    const py  = p => p * H;
    const sz  = s => Math.max(8, Math.round(s * W));
    const col = c => (c === 'variety' ? varietyColor : c);

    if (record.planting_date && C.planting_date)
      this._drawTextLegacy(ctx, record.planting_date, C.planting_date, px, py, sz, col);
    if (record.harvest_date && C.harvest_date)
      this._drawTextLegacy(ctx, record.harvest_date, C.harvest_date, px, py, sz, col);

    if (C.productivity) {
      const f = C.productivity;
      const v = record.productivity || '';
      const parts   = v.toString().split(',');
      const intPart = parts[0];
      const decPart = parts[1] !== undefined ? ',' + parts[1] : '';
      ctx.save();
      const bigSz   = sz(f.size);
      const smallSz = Math.round(bigSz * 0.50);
      ctx.fillStyle = col(f.color);
      if (decPart) {
        ctx.font = `${f.weight} ${bigSz}px "Helvetica Neue", Arial`;
        const intW = ctx.measureText(intPart).width;
        ctx.font = `${f.weight} ${smallSz}px "Helvetica Neue", Arial`;
        const decW = ctx.measureText(decPart).width;
        const startX = px(f.x) - (intW + decW) / 2;
        ctx.font = `${f.weight} ${bigSz}px "Helvetica Neue", Arial`; ctx.textAlign='left'; ctx.textBaseline='alphabetic';
        ctx.fillText(intPart, startX, py(f.y));
        ctx.font = `${f.weight} ${smallSz}px "Helvetica Neue", Arial`;
        ctx.fillText(decPart, startX + intW, py(f.y));
      } else {
        ctx.font = `${f.weight} ${bigSz}px "Helvetica Neue", Arial`; ctx.textAlign='center'; ctx.textBaseline='alphabetic';
        ctx.fillText(v, px(f.x), py(f.y));
      }
      ctx.restore();
    }

    if (record.unit && C.unit)
      this._drawTextLegacy(ctx, record.unit, C.unit, px, py, sz, col);
    if (C.location)
      this._drawTextLegacy(ctx, [record.city, record.state].filter(Boolean).join('/') || '-/-', C.location, px, py, sz, col, { shadow: true });
    if (record.producer_name && C.producer_name)
      this._drawTextLegacy(ctx, record.producer_name, C.producer_name, px, py, sz, col);
    if (record.farm_name && C.farm_name)
      this._drawTextLegacy(ctx, record.farm_name, C.farm_name, px, py, sz, col);
    if (record.area && C.area_badge) {
      const f = C.area_badge;
      const fSize = sz(f.size);
      ctx.save();
      ctx.font = `${f.weight} ${fSize}px "Helvetica Neue", Arial`;
      if (f.badge) {
        const metrics = ctx.measureText(record.area);
        const padX = fSize * 1.2; const padY = fSize * 0.5;
        const bW = metrics.width + padX * 2; const bH = fSize + padY * 2;
        const bX = px(f.x) - bW / 2; const bY = py(f.y) - fSize - padY + 2;
        ctx.fillStyle = varietyColor;
        this._roundRect(ctx, bX, bY, bW, bH, bH / 2); ctx.fill();
      }
      ctx.fillStyle = col(f.color); ctx.textAlign='center'; ctx.textBaseline='alphabetic';
      ctx.fillText(record.area, px(f.x), py(f.y));
      ctx.restore();
    }
  },

  _drawTextLegacy(ctx, text, f, px, py, sz, col, opts = {}) {
    ctx.save();
    ctx.font         = `${f.weight} ${sz(f.size)}px "Helvetica Neue", Arial`;
    ctx.fillStyle    = col(f.color);
    ctx.textAlign    = f.align;
    ctx.textBaseline = 'alphabetic';
    if (opts.shadow) { ctx.shadowColor='rgba(0,0,0,.45)'; ctx.shadowBlur=4; }
    else             { ctx.shadowColor='transparent'; ctx.shadowBlur=0; }
    ctx.fillText(text, px(f.x), py(f.y));
    ctx.restore();
  },

  /* ═══════════════════════════════════════════════════
     CALIBRADOR VISUAL (modo legado)
  ═══════════════════════════════════════════════════ */
  openCalibrator(variety, onSave) {
    if (!variety?.template_image) {
      App.Toast.show('Faça o upload do modelo antes de calibrar.', 'error');
      return;
    }
    const modal = document.getElementById('calibratorModal');
    if (!modal) return;
    const img = document.getElementById('calibratorImg');
    const existing = this._parseCoords(variety.field_coords);
    img.onload = () => this._initCalibrator(existing, onSave, variety.primary_color || '#3CB226');
    img.src = variety.template_image;
    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  },

  _calibFields: [
    { key: 'planting_date', label: 'Data de Plantio',    icon: '🌱', color: '#2E7D32' },
    { key: 'harvest_date',  label: 'Data de Colheita',   icon: '🚜', color: '#1565C0' },
    { key: 'productivity',  label: 'Produtividade (nº)', icon: '📊', color: '#E65100' },
    { key: 'unit',          label: 'Unidade (sc/ha)',     icon: '📏', color: '#7B1FA2' },
    { key: 'location',      label: 'Cidade/UF',           icon: '📍', color: '#C62828' },
    { key: 'producer_name', label: 'Nome do Produtor',    icon: '👤', color: '#37474F' },
    { key: 'farm_name',     label: 'Nome da Fazenda',     icon: '🏠', color: '#795548' },
    { key: 'area_badge',    label: 'Área Colhida',        icon: '📐', color: '#00695C' }
  ],

  _currentCoords: {}, _activeField: null, _calibOnSave: null, _calibColor: '#3CB226',

  _initCalibrator(existing, onSave, varietyColor) {
    this._currentCoords = JSON.parse(JSON.stringify(existing));
    this._calibOnSave = onSave; this._activeField = null; this._calibColor = varietyColor;
    const btnList = document.getElementById('calibFieldBtns');
    if (!btnList) return;
    btnList.innerHTML = this._calibFields.map(f => `
      <button class="calib-field-btn" data-field="${f.key}"
        onclick="CardGenerator._selectCalibField('${f.key}')" style="--calib-color:${f.color}">
        <span class="calib-field-icon">${f.icon}</span>
        <span class="calib-field-label">${f.label}</span>
        <span class="calib-field-status ${this._currentCoords[f.key] ? 'placed' : 'unset'}">
          ${this._currentCoords[f.key] ? '✓' : '–'}</span>
      </button>`).join('');
    this._renderCalibPins();
    const imgWrapper = document.getElementById('calibratorImgWrapper');
    if (imgWrapper) {
      const nw = imgWrapper.cloneNode(true);
      imgWrapper.parentNode?.replaceChild(nw, imgWrapper);
      nw.addEventListener('click', e => this._onCalibClick(e));
    }
    const hint = document.getElementById('calibHint');
    if (hint) hint.textContent = 'Selecione um campo à esquerda, depois clique na posição correta na imagem.';
  },

  _selectCalibField(key) {
    this._activeField = key;
    document.querySelectorAll('.calib-field-btn').forEach(b => b.classList.toggle('active', b.dataset.field === key));
    const f = this._calibFields.find(x => x.key === key);
    const hint = document.getElementById('calibHint');
    if (hint) hint.innerHTML = `<strong>Clique na imagem</strong> para posicionar: <em>${f?.label || key}</em>`;
  },

  _onCalibClick(e) {
    if (e.target.classList.contains('calib-pin')) return;
    if (!this._activeField) { App.Toast.show('Selecione um campo primeiro.', 'info'); return; }
    const wrapper = document.getElementById('calibratorImgWrapper');
    const rect    = wrapper.getBoundingClientRect();
    const xPct    = (e.clientX - rect.left) / rect.width;
    const yPct    = (e.clientY - rect.top)  / rect.height;
    this._currentCoords[this._activeField] = {
      ...(this.DEFAULT_COORDS[this._activeField] || {}),
      x: parseFloat(xPct.toFixed(4)),
      y: parseFloat(yPct.toFixed(4))
    };
    this._renderCalibPins(); this._refreshCalibBtns();
    const idx  = this._calibFields.findIndex(f => f.key === this._activeField);
    const next = this._calibFields[idx + 1];
    if (next) setTimeout(() => this._selectCalibField(next.key), 100);
    else {
      this._activeField = null;
      document.querySelectorAll('.calib-field-btn').forEach(b => b.classList.remove('active'));
      const hint = document.getElementById('calibHint');
      if (hint) hint.textContent = '✅ Todos os campos posicionados! Clique em "Salvar Calibração".';
    }
  },

  _renderCalibPins() {
    const wrapper = document.getElementById('calibratorImgWrapper');
    if (!wrapper) return;
    wrapper.querySelectorAll('.calib-pin').forEach(p => p.remove());
    Object.entries(this._currentCoords).forEach(([key, c]) => {
      if (!c) return;
      const f   = this._calibFields.find(x => x.key === key);
      const pin = document.createElement('div');
      pin.className        = 'calib-pin' + (key === this._activeField ? ' active-pin' : '');
      pin.style.left       = (c.x * 100) + '%';
      pin.style.top        = (c.y * 100) + '%';
      pin.style.background = f?.color || '#333';
      pin.title            = f?.label || key;
      pin.textContent      = f?.icon  || '●';
      pin.addEventListener('click', e => { e.stopPropagation(); this._selectCalibField(key); });
      wrapper.appendChild(pin);
    });
  },

  _refreshCalibBtns() {
    document.querySelectorAll('.calib-field-btn').forEach(btn => {
      const key = btn.dataset.field;
      const statusEl = btn.querySelector('.calib-field-status');
      if (!statusEl) return;
      const placed = !!this._currentCoords[key];
      statusEl.className   = 'calib-field-status ' + (placed ? 'placed' : 'unset');
      statusEl.textContent = placed ? '✓' : '–';
    });
  },

  saveCalibration() {
    if (this._calibOnSave) this._calibOnSave(JSON.parse(JSON.stringify(this._currentCoords)));
    this.closeCalibrator();
    App.Toast.show('Calibração salva com sucesso!', 'success');
  },

  closeCalibrator() {
    const modal = document.getElementById('calibratorModal');
    if (modal) modal.classList.add('hidden');
    document.body.style.overflow = '';
  },

  resetCalibration() {
    this._currentCoords = JSON.parse(JSON.stringify(this.DEFAULT_COORDS));
    this._renderCalibPins(); this._refreshCalibBtns();
    App.Toast.show('Coordenadas redefinidas para o padrão.', 'info');
  },

  /* ═══════════════════════════════════════════════════
     HELPERS
  ═══════════════════════════════════════════════════ */
  _parseCoords(jsonStr) {
    if (!jsonStr) return JSON.parse(JSON.stringify(this.DEFAULT_COORDS));
    try {
      const parsed = typeof jsonStr === 'string' ? JSON.parse(jsonStr) : jsonStr;
      const merged = JSON.parse(JSON.stringify(this.DEFAULT_COORDS));
      Object.entries(parsed).forEach(([k, v]) => {
        if (merged[k]) merged[k] = { ...merged[k], ...v };
        else merged[k] = v;
      });
      return merged;
    } catch { return JSON.parse(JSON.stringify(this.DEFAULT_COORDS)); }
  },

  async _getVariety(varietyId) {
    if (!varietyId) return null;
    if (this._varietyCache[varietyId]) return this._varietyCache[varietyId];
    try {
      const res = await API.getVarieties();
      (res.data || []).forEach(v => { this._varietyCache[v.id] = v; });
      return this._varietyCache[varietyId] || null;
    } catch { return null; }
  },

  _invalidateCache(varietyId) {
    if (varietyId) delete this._varietyCache[varietyId];
    else this._varietyCache = {};
  },

  _roundRect(ctx, x, y, w, h, r) {
    if (ctx.roundRect) { ctx.beginPath(); ctx.roundRect(x, y, w, h, r); return; }
    ctx.beginPath();
    ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r);
    ctx.lineTo(x + w, y + h - r); ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    ctx.lineTo(x + r, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - r);
    ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.closePath();
  },

  _esc(str) {
    if (!str) return '';
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  },

  _buildFilename(r, ext = '.png') {
    if (!r) return 'AgriCard_STINE' + ext;
    // Usa o nome padronizado do OneDrive se disponível
    if (r.card_filename) return r.card_filename;
    const safe = s => (s || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, '_').replace(/[^\w_-]/g, '');
    const date = new Date().toISOString().slice(0,10);
    const parts = [safe(r.producer_name), safe(r.variety_name), safe(r.city), date].filter(Boolean);
    return (parts.length ? parts.join('_') : 'AgriCard_STINE') + ext;
  },

  /**
   * Abre o preview de um card a partir do ID do registro
   * Usado pelo painel admin
   */
  async openPreviewById(recordId) {
    try {
      const record = await API.getRecord(recordId);
      const vRes   = await API.getVarieties();
      const v      = (vRes.data || []).find(x => x.id === record.variety_id);
      if (v) record._color = v.primary_color || '#2E7D32';
      await this.openPreview(record);
    } catch (err) {
      App.Toast.show('Erro ao carregar card.', 'error');
      console.error(err);
    }
  }

}; // fim CardGenerator
