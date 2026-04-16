/*
  Solum Core
  
  High-fidelity Spring Physics (Second-Order Dynamics).
  Simulates a mass-spring-damper system for interruptible, organic UI motion.
*/

#pragma once

namespace Core
{
    /**
     * @brief A robust spring solver using Semi-Implicit Euler integration.
     * Higher stability for UI animations compared to standard Euler.
     */
    struct Spring
    {
        float x = 0.0f;      // Current position
        float target = 0.0f; // Target position
        float v = 0.0f;      // Velocity
        
        // Physical parameters
        float stiffness = 180.0f; // k: Tension of the spring
        float damping = 22.0f;    // c: Friction (Critically damped at 2 * sqrt(stiffness) ~= 26)
        float mass = 1.0f;        // m: Inertia
        
        Spring() = default;
        Spring(float pos) : x(pos), target(pos) {}

        /**
         * @brief Updates the spring state.
         * @param dt Delta time in seconds.
         */
        void Update(float dt)
        {
            if (dt <= 0.0f) return;
            
            // a = F/m = (k*(target - x) - c*v) / m
            float force = stiffness * (target - x) - damping * v;
            float acceleration = force / mass;
            
            // Semi-Implicit Euler (More stable for springs)
            v += acceleration * dt;
            x += v * dt;
        }

        bool IsSettled(float eps = 0.001f, float velEps = 0.01f) const
        {
            return (std::abs(target - x) < eps) && (std::abs(v) < velEps);
        }
        
        void Reset(float pos)
        {
            x = pos;
            target = pos;
            v = 0.0f;
        }
    };
}
