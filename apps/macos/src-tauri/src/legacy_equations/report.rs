use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct FormulaResult {
    pub formula_id: String,
    pub status: FormulaStatus,
    #[serde(default)]
    pub warnings: Vec<String>,
    #[serde(default)]
    pub errors: Vec<String>,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "kebab-case")]
pub enum FormulaStatus {
    Replaced,
    Preserved,
    Skipped,
    Failed,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConversionReport {
    pub detected: u32,
    pub replaced: u32,
    pub preserved: u32,
    pub skipped: u32,
    pub failed: u32,
    pub source_unmodified: bool,
    pub input_sha256: String,
    pub output_sha256: String,
    #[serde(default)]
    pub package_valid: bool,
    #[serde(default)]
    pub formulas: Vec<FormulaResult>,
}

impl ConversionReport {
    pub fn validate(&self, expected_input_hash: &str) -> Result<(), String> {
        let accounted = self
            .replaced
            .checked_add(self.preserved)
            .and_then(|value| value.checked_add(self.skipped))
            .and_then(|value| value.checked_add(self.failed))
            .ok_or_else(|| "Conversion counts overflowed".to_string())?;
        if accounted != self.detected {
            return Err(format!(
                "Formula count conservation failed: detected={}, accounted={accounted}",
                self.detected
            ));
        }
        if !self.source_unmodified || self.input_sha256 != expected_input_hash {
            return Err("Worker reported a changed source document".to_string());
        }
        if !self.package_valid {
            return Err("Candidate package validation failed".to_string());
        }
        if self
            .formulas
            .iter()
            .any(|formula| formula.status == FormulaStatus::Replaced && !formula.errors.is_empty())
        {
            return Err("A formula with validation errors was marked replaced".to_string());
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn report() -> ConversionReport {
        ConversionReport {
            detected: 2,
            replaced: 1,
            preserved: 1,
            skipped: 0,
            failed: 0,
            source_unmodified: true,
            input_sha256: "a".repeat(64),
            output_sha256: "b".repeat(64),
            package_valid: true,
            formulas: vec![],
        }
    }

    #[test]
    fn rejects_count_conservation_failure() {
        let mut value = report();
        value.detected = 3;
        assert!(value.validate(&"a".repeat(64)).is_err());
    }

    #[test]
    fn rejects_source_hash_change() {
        assert!(report().validate(&"c".repeat(64)).is_err());
    }

    #[test]
    fn rejects_replaced_formula_with_validation_error() {
        let mut value = report();
        value.formulas.push(FormulaResult {
            formula_id: "F1".into(),
            status: FormulaStatus::Replaced,
            warnings: vec![],
            errors: vec!["token loss".into()],
        });
        assert!(value.validate(&"a".repeat(64)).is_err());
    }
}
