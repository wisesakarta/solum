/*
  Technical Standard
  
  Native animation controller providing cubic-bezier easing and hardware-accelerated 
  transitions for the premium UI/UX experience.
*/

#pragma once
#include <windows.h>
#include <cmath>

namespace Animation
{
    inline float EaseOutPunchy(float t) {
        float f = (t - 1.0f);
        return f * f * f * f * f + 1.0f;
    }

    struct Transition {
        float startValue;
        float endValue;
        float durationMs;
        float startTime;
        bool active;

        Transition() : startValue(0), endValue(0), durationMs(0), startTime(0), active(false) {}

        void Start(float start, float end, float duration) {
            startValue = start;
            endValue = end;
            durationMs = duration;
            startTime = static_cast<float>(GetTickCount64());
            active = true;
        }

        float GetCurrentValue() {
            if (!active) return endValue;
            float now = static_cast<float>(GetTickCount64());
            float elapsed = now - startTime;
            float t = elapsed / durationMs;
            if (t >= 1.0f) {
                active = false;
                return endValue;
            }
            return startValue + (endValue - startValue) * EaseOutPunchy(t);
        }
    };
}
