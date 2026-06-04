import React, { useState } from 'react';
import { Save, X } from 'lucide-react';
import type { Question, Response } from '../types';
import QuestionRenderer from './QuestionRenderer';
import { useQuestionLogic } from '../hooks/useQuestionLogic';

interface ResponseEditProps {
  response: Response;
  question: Question;
  onSave: (updatedAnswers: Record<string, any>) => void;  // Change this
  onCancel: () => void;
}

export default function ResponseEdit({
  response,
  question,
  onSave,
  onCancel,
}: ResponseEditProps) {
  const [answers, setAnswers] = useState(response.answers);
  const { getOrderedVisibleQuestions } = useQuestionLogic();

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    // Pass only the answers, not the full response
    onSave(answers);
  };

  // Get all questions from sections or fallback to followUpQuestions
  const allQuestions = question.sections.length > 0
    ? question.sections.flatMap(section => section.questions)
    : question.followUpQuestions;

  const visibleQuestions = getOrderedVisibleQuestions(allQuestions, answers);

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white dark:bg-gray-800 rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] overflow-y-auto m-4">
        <div className="sticky top-0 bg-white dark:bg-gray-800 border-b dark:border-gray-700 px-6 py-4 flex justify-between items-center">
          <h3 className="text-xl font-semibold text-gray-900 dark:text-white">
            Edit Response
          </h3>
          <div className="flex items-center space-x-3">
            <button
              type="button"
              onClick={onCancel}
              className="px-4 py-2 text-gray-700 bg-gray-100 rounded-lg hover:bg-gray-200 dark:bg-gray-700 dark:text-gray-300 dark:hover:bg-gray-600"
            >
              Cancel
            </button>
            <button
              type="submit"
              form="response-edit-form"
              className="flex items-center px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
            >
              <Save className="w-4 h-4 mr-2" />
              Save Changes
            </button>
            <button
              onClick={onCancel}
              className="text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-300"
            >
              <X className="w-6 h-6" />
            </button>
          </div>
        </div>

        <form id="response-edit-form" onSubmit={handleSubmit} className="p-6 space-y-6">
          {visibleQuestions.map((q) => (
            <QuestionRenderer
              key={q.id}
              question={q}
              value={answers[q.id]}
              trackingValue={answers[`${q.id}_tracking`]}
              onChange={(value) => setAnswers({ ...answers, [q.id]: value })}
              onTrackingChange={(value) => setAnswers({ ...answers, [`${q.id}_tracking`]: value })}
            />
          ))}
        </form>
      </div>
    </div>
  );
}