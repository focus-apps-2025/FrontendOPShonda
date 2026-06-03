// Updated CascadingFilterModal.tsx

import React, { useState, useMemo } from "react";
import { ChevronDown, X, Search } from "lucide-react";

interface FilterState {
  [questionId: string]: string[];
}

interface Question {
  id: string;
  text: string;
  type?: string;
  sectionId?: string;
  sectionTitle?: string;
}

interface Response {
  answers: Record<string, any>;
  submissionMetadata?: {
    location?: {
      city?: string;
      country?: string;
    };
  };
  timestamp?: string;
  createdAt?: string;
  [key: string]: any;
}

interface CascadingFilterModalProps {
  isOpen: boolean;
  onClose: () => void;
  questions: Question[]; // All questions from all sections
  responses: Response[];
  onApplyFilters: (filters: FilterState & { dates?: { startDate: string; endDate: string }; locations?: string[] }) => void;
}

export default function CascadingFilterModal({
  isOpen,
  onClose,
  questions,
  responses,
  onApplyFilters,
}: CascadingFilterModalProps) {
  const [filters, setFilters] = useState<FilterState>({});
  const [expandedQuestion, setExpandedQuestion] = useState<string | null>(null);
  const [searchTerms, setSearchTerms] = useState<Record<string, string>>({});
  const [dateRange, setDateRange] = useState({ startDate: "", endDate: "" });
  const [selectedLocations, setSelectedLocations] = useState<string[]>([]);
  const [expandedDateRange, setExpandedDateRange] = useState(false);
  const [expandedLocation, setExpandedLocation] = useState(false);
  const [locationSearchTerm, setLocationSearchTerm] = useState("");
  const [dropdownRect, setDropdownRect] = useState<{ top: number; left: number; width: number } | null>(null);
  const [expandedSection, setExpandedSection] = useState<string | null>(null);

  // Group questions by section
  const questionsBySection = useMemo(() => {
    const grouped: Record<string, { title: string; questions: Question[] }> = {};

    questions.forEach((question) => {
      const sectionId = question.sectionId || 'default';
      const sectionTitle = question.sectionTitle || 'General Questions';

      if (!grouped[sectionId]) {
        grouped[sectionId] = {
          title: sectionTitle,
          questions: []
        };
      }
      grouped[sectionId].questions.push(question);
    });

    return grouped;
  }, [questions]);

  const availableLocations = useMemo(() => {
    const locations = new Set<string>();
    responses.forEach((r) => {
      const meta = r.submissionMetadata?.location;
      if (meta) {
        const city = meta.city || "";
        const country = meta.country || "";
        const locationStr =
          city && country ? `${city}, ${country}` : country || "Unknown";
        if (locationStr !== "Unknown") {
          locations.add(locationStr);
        }
      }
    });
    return Array.from(locations).sort();
  }, [responses]);

  const filteredLocations = availableLocations.filter((loc) =>
    loc.toLowerCase().includes(locationSearchTerm.toLowerCase())
  );

  const checkAnswerMatch = (qId: string, selectedAnswers: string[], answer: any) => {
    if (!answer) return false;
    const question = questions.find(q => q.id === qId);

    if (question?.type === 'chassis-with-zone') {
      const zones = Array.isArray(answer.zone) ? answer.zone : [answer.zone];
      return zones.some((z: string) =>
        selectedAnswers.some(sel => String(z || '').toLowerCase() === String(sel || '').toLowerCase())
      );
    } else if (question?.type === 'chassis-without-zone') {
      return selectedAnswers.some(sel => String(answer.chassisNumber || '').toLowerCase() === String(sel || '').toLowerCase());
    }

    if (Array.isArray(answer)) {
      return answer.some((item) =>
        selectedAnswers.some(
          (sel) =>
            String(item).toLowerCase() === String(sel).toLowerCase()
        )
      );
    }
    return selectedAnswers.some(
      (sel) =>
        String(answer).toLowerCase() === String(sel).toLowerCase()
    );
  };

  const getAvailableAnswersForQuestion = (questionId: string): string[] => {
    const activeFilters = Object.entries(filters).filter(
      ([qId, answers]) => answers.length > 0 && qId !== questionId
    );

    let filteredResponses = responses;

    if (activeFilters.length > 0) {
      filteredResponses = responses.filter((response) =>
        activeFilters.every(([qId, selectedAnswers]) => {
          const answer = response.answers?.[qId];
          return checkAnswerMatch(qId, selectedAnswers, answer);
        })
      );
    }

    if (selectedLocations.length > 0) {
      filteredResponses = filteredResponses.filter((response) => {
        const meta = response.submissionMetadata?.location;
        if (!meta) return false;
        const city = meta.city || "";
        const country = meta.country || "";
        const locationStr =
          city && country ? `${city}, ${country}` : country || "Unknown";
        return selectedLocations.includes(locationStr);
      });
    }

    if (dateRange.startDate || dateRange.endDate) {
      filteredResponses = filteredResponses.filter((response) => {
        const timestamp = response.timestamp || response.createdAt;
        if (!timestamp) return false;
        const responseDate = new Date(timestamp).toISOString().split("T")[0];
        if (dateRange.startDate && dateRange.endDate) {
          return responseDate >= dateRange.startDate && responseDate <= dateRange.endDate;
        } else if (dateRange.startDate) {
          return responseDate >= dateRange.startDate;
        } else if (dateRange.endDate) {
          return responseDate <= dateRange.endDate;
        }
        return true;
      });
    }

    const answers = new Set<string>();
    const question = questions.find(q => q.id === questionId);

    filteredResponses.forEach((response) => {
      const answer = response.answers?.[questionId];
      if (answer !== null && answer !== undefined && answer !== "") {
        if (question?.type === 'chassis-with-zone') {
          const zones = Array.isArray(answer.zone) ? answer.zone : [answer.zone];
          zones.forEach((z: string) => {
            if (z) answers.add(String(z).trim());
          });
        } else if (question?.type === 'chassis-without-zone') {
          if (answer.chassisNumber) answers.add(String(answer.chassisNumber).trim());
        } else if (Array.isArray(answer)) {
          answer.forEach((a) => answers.add(String(a).trim()));
        } else {
          answers.add(String(answer).trim());
        }
      }
    });

    return Array.from(answers).sort();
  };

  const getAnswerCount = (
    questionId: string,
    answerValue: string
  ): number => {
    const activeFilters = Object.entries(filters).filter(
      ([qId, answers]) => answers.length > 0 && qId !== questionId
    );

    let filteredResponses = responses;

    if (activeFilters.length > 0) {
      filteredResponses = responses.filter((response) =>
        activeFilters.every(([qId, selectedAnswers]) => {
          const answer = response.answers?.[qId];
          return checkAnswerMatch(qId, selectedAnswers, answer);
        })
      );
    }

    if (selectedLocations.length > 0) {
      filteredResponses = filteredResponses.filter((response) => {
        const meta = response.submissionMetadata?.location;
        if (!meta) return false;
        const city = meta.city || "";
        const country = meta.country || "";
        const locationStr =
          city && country ? `${city}, ${country}` : country || "Unknown";
        return selectedLocations.includes(locationStr);
      });
    }

    if (dateRange.startDate || dateRange.endDate) {
      filteredResponses = filteredResponses.filter((response) => {
        const timestamp = response.timestamp || response.createdAt;
        if (!timestamp) return false;
        const responseDate = new Date(timestamp).toISOString().split("T")[0];
        if (dateRange.startDate && dateRange.endDate) {
          return responseDate >= dateRange.startDate && responseDate <= dateRange.endDate;
        } else if (dateRange.startDate) {
          return responseDate >= dateRange.startDate;
        } else if (dateRange.endDate) {
          return responseDate <= dateRange.endDate;
        }
        return true;
      });
    }

    const question = questions.find(q => q.id === questionId);
    return filteredResponses.filter((response) => {
      const answer = response.answers?.[questionId];
      if (!answer) return false;

      if (question?.type === 'chassis-with-zone') {
        const zones = Array.isArray(answer.zone) ? answer.zone : [answer.zone];
        return zones.some((z: string) => String(z || '').toLowerCase() === answerValue.toLowerCase());
      } else if (question?.type === 'chassis-without-zone') {
        return String(answer.chassisNumber || '').toLowerCase() === answerValue.toLowerCase();
      }

      if (Array.isArray(answer)) {
        return answer.some(
          (item) =>
            String(item).toLowerCase() === answerValue.toLowerCase()
        );
      }
      return String(answer).toLowerCase() === answerValue.toLowerCase();
    }).length;
  };

  const handleAnswerToggle = (questionId: string, answer: string) => {
    setFilters((prev) => {
      const currentAnswers = prev[questionId] || [];
      const updated = currentAnswers.includes(answer)
        ? currentAnswers.filter((a) => a !== answer)
        : [...currentAnswers, answer];

      return {
        ...prev,
        [questionId]: updated,
      };
    });
  };

  const handleLocationToggle = (location: string) => {
    setSelectedLocations((prev) =>
      prev.includes(location)
        ? prev.filter((l) => l !== location)
        : [...prev, location]
    );
  };

  const handleApply = () => {
    onApplyFilters({
      ...filters,
      dates: dateRange,
      locations: selectedLocations,
    });
    onClose();
  };

  const handleClearAll = () => {
    setFilters({});
    setSearchTerms({});
    setDateRange({ startDate: "", endDate: "" });
    setSelectedLocations([]);
    setLocationSearchTerm("");
  };

  const appliedFiltersCount =
    Object.values(filters).reduce((sum, answers) => sum + answers.length, 0) +
    (dateRange.startDate || dateRange.endDate ? 1 : 0) +
    (selectedLocations.length > 0 ? 1 : 0);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-lg w-full max-w-6xl max-h-[90vh] overflow-hidden flex flex-col relative">
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-300 flex items-center justify-between relative z-40">
          <h2 className="text-lg font-bold text-gray-900">
            Question Filters{" "}
            {appliedFiltersCount > 0 && `(${appliedFiltersCount} selected)`}
          </h2>
          <button onClick={onClose} className="text-gray-600 hover:text-gray-900">
            <X className="w-6 h-6" />
          </button>
        </div>

        {/* Main Content */}
        <div className="flex-1 overflow-y-auto p-6 flex flex-col relative z-20">
          {/* Date Range & Location - Side by Side */}
          <div className="grid grid-cols-2 gap-4 mb-6">
            {/* Date Range Dropdown */}
            <div className="relative">
              <button
                onClick={() => {
                  setExpandedDateRange(!expandedDateRange);
                  setExpandedLocation(false);
                  setExpandedQuestion(null);
                  setExpandedSection(null);
                }}
                className="w-full px-4 py-3 bg-gray-50 hover:bg-gray-100 border border-gray-300 rounded-lg flex items-center justify-between transition-colors"
              >
                <div className="flex-1 text-left flex items-center gap-2">
                  <p className="font-semibold text-gray-900 text-sm">Date Range</p>
                  {(dateRange.startDate || dateRange.endDate) && (
                    <span className="inline-flex items-center justify-center w-5 h-5 bg-indigo-600 text-white text-xs font-medium rounded-full">
                      1
                    </span>
                  )}
                </div>
                <ChevronDown
                  className={`w-5 h-5 text-gray-600 transition-transform flex-shrink-0 ${expandedDateRange ? "transform rotate-180" : ""
                    }`}
                />
              </button>

              {expandedDateRange && (
                <div className="absolute top-full left-0 right-0 mt-2 bg-white border border-gray-300 rounded-lg p-4 z-50 shadow-lg">
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-medium text-gray-700 mb-2">
                        Start Date
                      </label>
                      <input
                        type="date"
                        value={dateRange.startDate}
                        onChange={(e) =>
                          setDateRange({ ...dateRange, startDate: e.target.value })
                        }
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-gray-700 mb-2">
                        End Date
                      </label>
                      <input
                        type="date"
                        value={dateRange.endDate}
                        onChange={(e) =>
                          setDateRange({ ...dateRange, endDate: e.target.value })
                        }
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
                      />
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Location Dropdown */}
            <div className="relative">
              <button
                onClick={() => {
                  setExpandedLocation(!expandedLocation);
                  setExpandedDateRange(false);
                  setExpandedQuestion(null);
                  setExpandedSection(null);
                }}
                className="w-full px-4 py-3 bg-gray-50 hover:bg-gray-100 border border-gray-300 rounded-lg flex items-center justify-between transition-colors"
              >
                <div className="flex-1 text-left flex items-center gap-2">
                  <p className="font-semibold text-gray-900 text-sm">Locations</p>
                  {selectedLocations.length > 0 && (
                    <span className="inline-flex items-center justify-center w-5 h-5 bg-indigo-600 text-white text-xs font-medium rounded-full">
                      {selectedLocations.length}
                    </span>
                  )}
                </div>
                <ChevronDown
                  className={`w-5 h-5 text-gray-600 transition-transform flex-shrink-0 ${expandedLocation ? "transform rotate-180" : ""
                    }`}
                />
              </button>

              {expandedLocation && (
                <div className="absolute top-full right-0 mt-2 bg-white border border-gray-300 rounded-lg p-4 z-50 shadow-lg w-full">
                  <div className="mb-3 relative">
                    <Search className="absolute left-2.5 top-2.5 w-4 h-4 text-gray-400" />
                    <input
                      type="text"
                      placeholder="Search locations..."
                      value={locationSearchTerm}
                      onChange={(e) => setLocationSearchTerm(e.target.value)}
                      className="w-full pl-8 pr-3 py-2 border border-gray-300 rounded text-xs focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
                    />
                  </div>

                  <div className="space-y-2 max-h-48 overflow-y-auto">
                    {filteredLocations.length === 0 ? (
                      <p className="text-xs text-gray-500 text-center py-3">
                        No locations found
                      </p>
                    ) : (
                      filteredLocations.map((location) => (
                        <label
                          key={location}
                          className="flex items-center gap-2 cursor-pointer p-2 hover:bg-gray-50 rounded transition-colors"
                        >
                          <input
                            type="checkbox"
                            checked={selectedLocations.includes(location)}
                            onChange={() => handleLocationToggle(location)}
                            className="w-4 h-4 rounded border-gray-300 text-indigo-600 cursor-pointer"
                          />
                          <span className="text-sm text-gray-900 truncate">
                            {location}
                          </span>
                        </label>
                      ))
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Questions by Section - Accordion Style */}
          <div className="flex-1 flex flex-col overflow-hidden">
            <div className="flex items-center justify-between mb-4 flex-shrink-0">
              <h3 className="font-semibold text-gray-900">Questions by Section</h3>
              {Object.values(filters).reduce((sum, answers) => sum + answers.length, 0) > 0 && (
                <span className="text-xs text-indigo-600 font-medium">
                  {Object.values(filters).reduce((sum, answers) => sum + answers.length, 0)} selected
                </span>
              )}
            </div>

            <div className="flex-1 pr-2 relative" style={{ overflowY: "auto" }} onScroll={() => setExpandedQuestion(null)}>
              {Object.keys(questionsBySection).length === 0 ? (
                <div className="text-center py-8 text-gray-500">
                  No questions available
                </div>
              ) : (
                <div className="space-y-4">
                  {Object.entries(questionsBySection).map(([sectionId, { title, questions: sectionQuestions }]) => (
                    <div key={sectionId} className="border border-gray-200 rounded-lg overflow-hidden">
                      {/* Section Header */}
                      <button
                        onClick={() => setExpandedSection(expandedSection === sectionId ? null : sectionId)}
                        className="w-full px-4 py-3 bg-gradient-to-r from-indigo-50 to-blue-50 hover:from-indigo-100 hover:to-blue-100 border-b border-gray-200 flex items-center justify-between transition-colors"
                      >
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-indigo-700 text-sm uppercase tracking-wide">
                            {title}
                          </span>
                          <span className="text-xs text-gray-500 bg-white/50 px-2 py-0.5 rounded-full">
                            {sectionQuestions.length} questions
                          </span>
                        </div>
                        <ChevronDown
                          className={`w-4 h-4 text-indigo-600 transition-transform ${expandedSection === sectionId ? "transform rotate-180" : ""
                            }`}
                        />
                      </button>

                      {/* Section Questions Grid */}
                      {expandedSection === sectionId && (
                        <div className="p-4 bg-white">
                          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                            {sectionQuestions.map((question) => {
                              const isExpanded = expandedQuestion === question.id;
                              const selectedAnswers = filters[question.id] || [];

                              return (
                                <div key={question.id} className="relative">
                                  <button
                                    onClick={(e) => {
                                      if (expandedQuestion === question.id) {
                                        setExpandedQuestion(null);
                                      } else {
                                        const rect = e.currentTarget.getBoundingClientRect();
                                        setDropdownRect({
                                          top: rect.bottom,
                                          left: rect.left,
                                          width: rect.width
                                        });
                                        setExpandedQuestion(question.id);
                                        setExpandedDateRange(false);
                                        setExpandedLocation(false);
                                        setExpandedSection(null);
                                      }
                                    }}
                                    className="w-full px-4 py-3 bg-gray-50 hover:bg-gray-100 border border-gray-300 rounded-lg flex items-center justify-between transition-colors"
                                  >
                                    <div className="flex-1 text-left min-w-0 flex items-center gap-2">
                                      <p className="font-semibold text-gray-900 text-xs line-clamp-2 flex-1">
                                        {question.text || "Unnamed Question"}
                                      </p>
                                      {selectedAnswers.length > 0 && (
                                        <span className="flex-shrink-0 inline-flex items-center justify-center w-5 h-5 bg-indigo-600 text-white text-xs font-medium rounded-full">
                                          {selectedAnswers.length}
                                        </span>
                                      )}
                                    </div>
                                    <ChevronDown
                                      className={`w-4 h-4 text-gray-600 transition-transform flex-shrink-0 ml-2 ${isExpanded ? "transform rotate-180" : ""
                                        }`}
                                    />
                                  </button>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="bg-gray-100 px-6 py-4 border-t border-gray-300 flex items-center justify-between gap-3 relative z-30">
          <button
            onClick={handleClearAll}
            className="px-4 py-2 text-gray-900 bg-white border border-gray-400 hover:bg-gray-50 text-sm font-medium rounded transition-colors"
          >
            Clear All
          </button>
          <div className="flex gap-3">
            <button
              onClick={onClose}
              className="px-4 py-2 text-gray-900 bg-white border border-gray-400 hover:bg-gray-50 text-sm font-medium rounded transition-colors"
            >
              Cancel
            </button>
            <button
              onClick={handleApply}
              className="px-4 py-2 bg-indigo-600 text-white hover:bg-indigo-700 text-sm font-medium rounded transition-colors"
            >
              Apply Filters
            </button>
          </div>
        </div>

        {/* Fixed Dropdown for Questions */}
        {expandedQuestion && dropdownRect && (() => {
          const question = questions.find(q => q.id === expandedQuestion);
          if (!question) return null;

          const selectedAnswers = filters[question.id] || [];
          const searchTerm = searchTerms[question.id] || "";
          const availableAnswers = getAvailableAnswersForQuestion(question.id);
          const filteredAnswers = availableAnswers.filter((answer) =>
            answer.toLowerCase().includes(searchTerm.toLowerCase())
          );

          return (
            <>
              <div className="fixed inset-0 z-[55]" onClick={() => setExpandedQuestion(null)} />
              <div
                className="fixed bg-white border border-gray-300 rounded-lg p-3 z-[60] shadow-xl flex flex-col"
                style={{
                  top: dropdownRect.top + 4,
                  left: dropdownRect.left,
                  width: dropdownRect.width,
                  maxHeight: '300px'
                }}
              >
                <div className="mb-3 relative flex-shrink-0">
                  <Search className="absolute left-2.5 top-2.5 w-4 h-4 text-gray-400" />
                  <input
                    type="text"
                    placeholder="Search..."
                    value={searchTerm}
                    onChange={(e) =>
                      setSearchTerms({
                        ...searchTerms,
                        [question.id]: e.target.value,
                      })
                    }
                    className="w-full pl-8 pr-3 py-2 border border-gray-300 rounded text-xs focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
                    autoFocus
                  />
                </div>

                <div className="flex gap-2 mb-3 flex-shrink-0">
                  <button
                    onClick={() => {
                      setFilters((prev) => ({
                        ...prev,
                        [question.id]: availableAnswers,
                      }));
                    }}
                    className="flex-1 px-4 py-2.5 text-[10px] font-black uppercase tracking-widest text-white bg-indigo-600 hover:bg-indigo-700 rounded-xl transition-all shadow-md shadow-indigo-600/10 active:scale-95"
                  >
                    Select All
                  </button>
                  <button
                    onClick={() => {
                      setFilters((prev) => ({
                        ...prev,
                        [question.id]: [],
                      }));
                    }}
                    className="flex-1 px-4 py-2.5 text-[10px] font-black uppercase tracking-widest text-gray-500 bg-gray-100 hover:bg-gray-200 rounded-xl transition-all active:scale-95"
                  >
                    Clear
                  </button>
                </div>

                <div className="space-y-1 overflow-y-auto flex-1 custom-scrollbar pr-1">
                  {filteredAnswers.length === 0 ? (
                    <div className="flex flex-col items-center justify-center py-8 text-gray-400">
                      <Search className="w-8 h-8 opacity-20 mb-2" />
                      <p className="text-[10px] font-bold uppercase tracking-widest">No results</p>
                    </div>
                  ) : (
                    filteredAnswers.map((answer) => {
                      const count = getAnswerCount(question.id, answer);
                      const isSelected = selectedAnswers.includes(answer);

                      return (
                        <label
                          key={answer}
                          className={`flex items-center gap-3 p-3 rounded-xl cursor-pointer transition-all duration-200 group/label ${isSelected
                            ? 'bg-indigo-600 text-white shadow-lg shadow-indigo-600/20 translate-x-1'
                            : 'hover:bg-indigo-50 text-gray-700 hover:text-indigo-700'
                            }`}
                        >
                          <div className={`w-4 h-4 rounded border flex items-center justify-center transition-colors ${isSelected ? 'bg-white border-white' : 'bg-white border-gray-300 group-hover/label:border-indigo-500'
                            }`}>
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={() => handleAnswerToggle(question.id, answer)}
                              className="sr-only"
                            />
                            {isSelected && (
                              <div className="w-2.5 h-2.5 bg-indigo-600 rounded-[2px]" />
                            )}
                          </div>

                          <div className="flex-1 min-w-0">
                            <p className="text-xs font-bold truncate">
                              {answer}
                            </p>
                          </div>

                          <span className={`text-[10px] font-black tracking-widest px-2 py-0.5 rounded-full ${isSelected ? 'bg-white/20 text-white' : 'bg-gray-100 text-gray-500 group-hover/label:bg-indigo-100 group-hover/label:text-indigo-600'
                            }`}>
                            {count}
                          </span>
                        </label>
                      );
                    })
                  )}
                </div>
              </div>
            </>
          );
        })()}
      </div>
    </div>
  );
}