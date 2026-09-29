from dataclasses import replace

import pytest

from excel_catalog_pipeline.context_profiles import load_prompt, prompt_digest
from excel_catalog_pipeline.context_source import ContextError
from excel_catalog_pipeline.models import ContextConfig


def test_prompt_profiles_language_and_content_hash():
    ja = ContextConfig(prompt_profile="default-ja")
    en = replace(ja, prompt_profile="default-en")
    assert "Japanese" in load_prompt(ja)
    assert "English" in load_prompt(en)
    assert "Relationships between sheets" in load_prompt(en, stage="workbook")
    assert prompt_digest(ja) != prompt_digest(en)
    with pytest.raises(ContextError, match="Unknown context profile"):
        load_prompt(replace(ja, prompt_profile="../unknown"))
