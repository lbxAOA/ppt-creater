function scoreFamily(family, request) {
  const requestedUseCase = (request.use_case || '').toLowerCase();
  const requestedTags = new Set((request.visual_tags || []).map((tag) => tag.toLowerCase()));
  const useCaseScore = (family.use_cases || []).some((value) => value.toLowerCase() === requestedUseCase) ? 8 : 0;
  const tagScore = (family.visual_tags || []).reduce(
    (score, tag) => score + (requestedTags.has(tag.toLowerCase()) ? 2 : 0), 0
  );
  const ratioScore = !request.aspect_ratio || family.aspect_ratio === request.aspect_ratio ? 2 : -5;
  return useCaseScore + tagScore + ratioScore;
}

function chooseFamily(families, request) {
  if (!Array.isArray(families) || families.length === 0) {
    throw new Error('No template families are registered.');
  }
  const ranked = families
    .map((family) => ({ family, score: scoreFamily(family, request) }))
    .sort((a, b) => b.score - a.score);
  if (ranked[0].score < 1) {
    throw new Error('No template family matches the requested use case and visual constraints.');
  }
  return ranked[0].family;
}

function buildDeckPlan(families, request) {
  const family = request.family_id
    ? families.find((item) => item.family_id === request.family_id)
    : chooseFamily(families, request);
  if (!family) throw new Error(`Unknown template family: ${request.family_id}`);

  const allowedAnimations = new Set((family.animation_policy?.allowed_patterns || []).concat('none'));
  const slides = (request.slides || []).map((requestedSlide, index) => {
    const template = family.slides.find((slide) => slide.type === requestedSlide.type);
    if (!template) throw new Error(`Family ${family.family_id} has no slide type: ${requestedSlide.type}`);
    const animation = requestedSlide.animation || template.animation_pattern || 'none';
    if (!allowedAnimations.has(animation)) {
      throw new Error(`Animation ${animation} is not allowed by family ${family.family_id}`);
    }
    if (template.animation_pattern !== 'none' && animation !== template.animation_pattern) {
      throw new Error(`Template slide ${template.slide_id} preserves ${template.animation_pattern}; requested ${animation}.`);
    }
    const content = requestedSlide.content || {};
    for (const slot of template.slots || []) {
      if (!(slot in content) || content[slot] === '') {
        throw new Error(`Slide ${index + 1} (${requestedSlide.type}) is missing required slot: ${slot}`);
      }
    }
    return {
      index: index + 1,
      type: requestedSlide.type,
      template_slide_id: template.slide_id,
      source_slide_index: template.source_slide_index,
      animation_pattern: animation,
      content
    };
  });
  return {
    version: 1,
    title: request.title || '',
    family_id: family.family_id,
    aspect_ratio: family.aspect_ratio,
    theme: family.theme,
    slides
  };
}

module.exports = { chooseFamily, buildDeckPlan, scoreFamily };
